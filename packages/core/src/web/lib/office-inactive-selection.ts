/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A kept text selection that LOOKS selected, for the plain text boxes the
 * three editors type in (a shape's words, a slide text box, a title, a cell).
 *
 * A browser stops painting a `<textarea>` / `<input>` selection the moment the
 * box loses focus (Safari keeps only a faint gray, invisible on a coloured
 * shape), while the box -- and the deck's own saved range -- keep it: the next
 * ribbon command still acts on it. Clicking the ribbon therefore read as "my
 * selection was cancelled" although it was not.
 *
 * The rule, the same for every such box and bound to what the commands use:
 *
 * - Leaving the box FOR editor chrome (the ribbon, a menu, a task pane, a
 *   dialog) keeps the range and draws it the way a live selection is drawn,
 *   over the box, for as long as it is kept.
 * - Leaving it for anything else drops the range: the box collapses it and a
 *   bubbling `office-inactive-selection-collapse` event tells the host, so a
 *   host that saved the range for its commands (the deck) drops it too.
 *
 * "For chrome" is decided by what the pointer PRESSED, not by what took the
 * focus: Safari gives a clicked button no focus at all, so the focus lands on
 * the page body even when the click was on the ribbon.
 */

import {
  clearOfficeTextRangeHighlight,
  officeTextRangeHighlightsSupported,
  setOfficeTextRangeHighlight,
  type OfficeTextRangeHighlightPaint,
} from "./office-text-range-highlight";

/** Editor chrome that legitimately takes the pointer or the focus while acting on the selection. */
export const OFFICE_CHROME_SELECTOR = [
  ".note-editor-toolbar-stack",
  ".note-editor-ribbon-tabs",
  ".note-editor-ribbon-strip",
  ".note-editor-ribbon-body",
  ".note-editor-ribbon-track",
  ".note-editor-office-toolbar",
  ".floating-toolbar",
  ".ai-selection-menu",
  ".ai-panel",
  "[data-office-ribbon]",
  "[data-office-chrome]",
  "[role='toolbar']",
  "[role='menu']",
  "[role='menubar']",
  "[role='listbox']",
  "[data-radix-popper-content-wrapper]",
  "[data-radix-portal]",
  ".office-dialog",
  "[role='dialog']",
  "[role='alertdialog']",
].join(", ");

export const OFFICE_INACTIVE_SELECTION_ATTRIBUTE = "data-office-inactive-selection";
/**
 * On a box whose host paints the box's selection itself, over glyphs the
 * box's own layout does not match (a PDF line drawn from the page's glyph
 * quads): the shared ghost is not painted for it, the host keeps its own
 * band (`watchOfficeInactiveSelection`).
 */
export const OFFICE_INACTIVE_SELECTION_SELF_PAINTED_ATTRIBUTE = "data-office-inactive-selection-self";
/** The CSS Custom Highlight a held rich box's range is painted with. */
export const OFFICE_HELD_RANGE_HIGHLIGHT = "office-held-range";
// Painted through the shared registry (its rule is added from script: the
// dev bundler's CSS parser drops `::highlight()`), one owner per held box,
// so a box released in one window no longer clears another's held range.
const OFFICE_HELD_RANGE_PAINT: OfficeTextRangeHighlightPaint = {
  name: OFFICE_HELD_RANGE_HIGHLIGHT,
  css: "background-color: color-mix(in srgb, Highlight 70%, transparent);",
};

function toStaticRange(range: AbstractRange): StaticRange {
  return new StaticRange({
    startContainer: range.startContainer,
    startOffset: range.startOffset,
    endContainer: range.endContainer,
    endOffset: range.endOffset,
  });
}
/** Fired (bubbling) at a box whose kept range was dropped. */
export const OFFICE_INACTIVE_SELECTION_COLLAPSE_EVENT = "office-inactive-selection-collapse";

type TextBox = HTMLTextAreaElement | HTMLInputElement;

const TEXT_INPUT_TYPES = new Set(["text", "search", "url", "tel", "email", ""]);
/** A press this recent is what moved the focus. */
const PRESS_WINDOW_MS = 1_500;

let lastPress: { target: EventTarget | null; at: number } | null = null;
let pressTrackerInstalled = false;

function trackPresses(root: Document): void {
  if (pressTrackerInstalled) return;
  pressTrackerInstalled = true;
  const remember = (event: Event) => {
    lastPress = { target: event.target, at: Date.now() };
  };
  root.addEventListener("pointerdown", remember, true);
  root.addEventListener("mousedown", remember, true);
}

/** The element the pointer pressed just now, or null. */
export function officeRecentPressTarget(): Element | null {
  if (!lastPress || Date.now() - lastPress.at > PRESS_WINDOW_MS) return null;
  return lastPress.target instanceof Element ? lastPress.target : null;
}

function inChrome(node: EventTarget | Element | null, selector = OFFICE_CHROME_SELECTOR): boolean {
  return node instanceof Element && Boolean(node.closest(selector));
}

/**
 * Whether the focus moved to another place the user types text in, outside
 * the chrome (another cell, another text box). That is a selection of its
 * own whatever was pressed before: a jump from a task pane (the Watch
 * Window, a Go To list) that puts the keyboard on another cell leaves the
 * old range behind, though the press that started it was in the chrome
 * (live audit S-141: H9's text stayed painted selected once the active
 * cell was J8).
 */
function focusMovedToAnotherTextEntry(node: EventTarget | null): boolean {
  if (!(node instanceof HTMLElement) || inChrome(node)) return false;
  if (node instanceof HTMLTextAreaElement) return true;
  if (node instanceof HTMLInputElement) return TEXT_INPUT_TYPES.has(node.type);
  return node.isContentEditable;
}

/**
 * Marks a control that formats characters (font, size, colour, Bold...).
 * The shared font picker, value box, colour palette and Home commands put it
 * on themselves when they format text, so an editor can tell a press there
 * (which acts on the characters selected) from any other ribbon command.
 */
export const OFFICE_TEXT_FORMAT_CONTROL_ATTRIBUTE = "data-office-text-format-control";

/** Whether `node` is (inside) a character-formatting control. */
export function isOfficeTextFormatControl(node: EventTarget | null): boolean {
  return inChrome(node, `[${OFFICE_TEXT_FORMAT_CONTROL_ATTRIBUTE}]`);
}

/**
 * After a character format chosen from a menu (a font list closes once it
 * is done and leaves the focus on the page), the text box being edited gets
 * the keyboard back at its selection, as Office returns to the text. A focus
 * the user has since put anywhere else is left alone.
 */
export function officeReturnFocusAfterTextFormat(
  getBox: () => (HTMLElement & {
    selectionStart?: number | null;
    selectionEnd?: number | null;
    setSelectionRange?: (start: number, end: number) => void;
  }) | null,
  range?: { start: number; end: number },
  /** Puts the range back in a box with no selection API of its own (a
   * contenteditable body); a field's own `setSelectionRange` otherwise. */
  select?: (box: HTMLElement, start: number, end: number) => void,
): void {
  if (typeof window === "undefined") return;
  const settle = () => {
    const box = getBox();
    if (!box?.isConnected) return;
    const active = box.ownerDocument.activeElement;
    if (active && box.contains(active)) return;
    if (active && active !== box.ownerDocument.body && !isOfficeTextFormatControl(active)) return;
    const start = range?.start ?? box.selectionStart ?? 0;
    const end = range?.end ?? box.selectionEnd ?? start;
    box.focus({ preventScroll: true });
    try {
      if (box.setSelectionRange) box.setSelectionRange(start, end);
      else select?.(box, start, end);
    } catch {
      // A box whose type has no selection API keeps the caret it gets.
    }
  };
  window.requestAnimationFrame(() => window.requestAnimationFrame(settle));
}

/** Whether a blur went into a character-formatting control (focus or press). */
export function officeBlurWentIntoTextFormatControl(related: EventTarget | null): boolean {
  return isOfficeTextFormatControl(related) || isOfficeTextFormatControl(officeRecentPressTarget());
}

/**
 * Chrome a cell entry stays open for, as Excel keeps edit mode across it: a
 * character-formatting control (it acts on the characters selected), the
 * ribbon's tab strip (switching tabs changes nothing), and inert chrome --
 * a group label, a gap, a pane's frame -- that is no command at all. Any
 * other command commits the entry first, as it acts on the cell.
 */
export function officeChromeKeepsCellEntry(target: EventTarget | null): boolean {
  if (isOfficeTextFormatControl(target)) return true;
  if (!(target instanceof Element)) return false;
  if (target.closest(".note-editor-ribbon-tabs, [role='tablist'], [role='tab']")) return true;
  if (!inChrome(target)) return false;
  return !target.closest(
    "button, [role='button'], [role='menuitem'], [role='menuitemcheckbox'], [role='menuitemradio'], [role='option'], [role='checkbox'], [role='radio'], [role='switch'], [role='slider'], [role='combobox'], input, select, textarea, a[href], [contenteditable='true']",
  );
}

/** Whether a blur went into chrome a cell entry stays open for (focus or press). */
export function officeBlurKeepsCellEntry(related: EventTarget | null): boolean {
  const pressed = officeRecentPressTarget();
  if (pressed) return officeChromeKeepsCellEntry(pressed);
  return officeChromeKeepsCellEntry(related);
}

/**
 * Whether a blur went INTO editor chrome: the focus moved there, or the
 * pointer pressed there (the focus may have gone nowhere). `selector` lets a
 * host with its own chrome list (the document editor) ask the same question.
 */
export function officeBlurWentIntoChrome(
  related: EventTarget | null,
  selector = OFFICE_CHROME_SELECTOR,
): boolean {
  return inChrome(related, selector) || inChrome(officeRecentPressTarget(), selector);
}

function isTextBox(target: EventTarget | null): target is TextBox {
  if (target instanceof HTMLTextAreaElement) return true;
  return target instanceof HTMLInputElement && TEXT_INPUT_TYPES.has(target.type);
}

function hasRange(box: TextBox): boolean {
  const start = box.selectionStart;
  const end = box.selectionEnd;
  return typeof start === "number" && typeof end === "number" && end > start;
}

/** A box that is gone or no longer drawn has no range to show. */
function boxShown(box: TextBox): boolean {
  if (!box.isConnected || box.hidden) return false;
  const style = box.ownerDocument.defaultView?.getComputedStyle(box);
  return !style || (style.display !== "none" && style.visibility !== "hidden");
}

const MIRRORED_PROPERTIES = [
  "box-sizing",
  "padding-top",
  "padding-right",
  "padding-bottom",
  "padding-left",
  "border-top-width",
  "border-right-width",
  "border-bottom-width",
  "border-left-width",
  "font-family",
  "font-size",
  "font-weight",
  "font-style",
  "font-variant",
  "font-stretch",
  "line-height",
  "letter-spacing",
  "word-spacing",
  "text-align",
  "text-indent",
  "text-transform",
  "direction",
  "white-space",
  "overflow-wrap",
  "word-break",
  "tab-size",
  // The same glyph advances: kerning, ligatures, numeric and variation
  // settings move every later character (a band that drifted off the end of
  // an amount in a font with tabular figures or kerning pairs).
  "font-kerning",
  "font-feature-settings",
  "font-variation-settings",
  "font-variant-ligatures",
  "font-variant-numeric",
  "font-variant-east-asian",
  "font-optical-sizing",
  "text-rendering",
  "writing-mode",
  "text-orientation",
  "unicode-bidi",
] as const;

/** What the ghost was drawn from; a change means it is redrawn. */
function ghostKey(box: TextBox): string {
  const rect = box.getBoundingClientRect();
  const computed = box.ownerDocument.defaultView?.getComputedStyle(box);
  return [
    box.value,
    box.selectionStart,
    box.selectionEnd,
    // Exact, not rounded: a band a fraction of a pixel off its words still
    // reads as off them on a dense screen.
    rect.left,
    rect.top,
    rect.width,
    rect.height,
    box.scrollTop,
    box.scrollLeft,
    // Everything the mirror lays its words out with: a ribbon command that
    // changes the size, the line height, the padding or the spacing moves
    // the words inside an unmoved box.
    ...(computed ? MIRRORED_PROPERTIES.map((property) => computed.getPropertyValue(property)) : []),
    computed?.transform ?? "",
  ].join("|");
}

/**
 * A mirror of the box over it that draws the kept range as a selection band
 * in the system selection colour, the words themselves showing through.
 *
 * The mirror is the box's next sibling, laid over it at the box's own place
 * and size inside the box's own stacking context: it takes the same scale a
 * slide canvas draws the box with, and every piece of chrome the range is
 * kept for (a context menu, a ribbon dropdown, a task pane, a dialog) opens
 * above it, as it opens above the box. A box without a positioned ancestor
 * (none in the editors; a bare page) gets the mirror fixed over it instead,
 * laid out at the box's own size and scaled to the drawn size.
 */
/** A run of a text box's characters painted over its words. */
export interface OfficeTextControlMark {
  readonly start: number;
  readonly end: number;
  /** The run's background (the selection colour when absent). */
  readonly background?: string;
  /** Extra attributes on the run, for its owner (a find hit's markers). */
  readonly attributes?: Readonly<Record<string, string>>;
}

function paintGhost(
  box: TextBox,
  root: Document,
  marks?: readonly OfficeTextControlMark[],
): HTMLElement | null {
  const rect = box.getBoundingClientRect();
  if (rect.width <= 0 || rect.height <= 0) return null;
  const computed = root.defaultView?.getComputedStyle(box);
  if (!computed) return null;
  const ghost = root.createElement("div");
  ghost.setAttribute(OFFICE_INACTIVE_SELECTION_ATTRIBUTE, "");
  ghost.setAttribute("aria-hidden", "true");
  const style = ghost.style;
  const layoutWidth = box.offsetWidth > 0 ? box.offsetWidth : rect.width;
  const layoutHeight = box.offsetHeight > 0 ? box.offsetHeight : rect.height;
  const beside = box.offsetParent !== null && box.parentElement !== null;
  if (beside) {
    style.position = "absolute";
    style.left = `${box.offsetLeft}px`;
    style.top = `${box.offsetTop}px`;
    // Over the box, under whatever the host stacks above the box itself.
    style.zIndex = computed.position !== "static" && computed.zIndex !== "auto" ? computed.zIndex : "auto";
    // offsetLeft/offsetTop are the box's place BEFORE its own transform. A
    // box centred in its shape (`top: 50%; translateY(-50%)`) is drawn half
    // its height higher than that, so the band sat below the words it marks
    // and read as a second, shifted copy of them. Same place, same size,
    // same transform: the band lands on the box.
    if (computed.transform && computed.transform !== "none") {
      style.transform = computed.transform;
      style.transformOrigin = computed.transformOrigin;
    }
    for (const property of ["translate", "rotate", "scale"] as const) {
      const value = computed.getPropertyValue(property);
      if (value && value !== "none") style.setProperty(property, value);
    }
  } else {
    style.position = "fixed";
    style.left = `${rect.left}px`;
    style.top = `${rect.top}px`;
    style.transformOrigin = "0 0";
    style.transform = `scale(${rect.width / layoutWidth}, ${rect.height / layoutHeight})`;
    style.zIndex = "2147483000";
  }
  style.width = `${layoutWidth}px`;
  style.height = `${layoutHeight}px`;
  style.margin = "0";
  style.overflow = "hidden";
  style.pointerEvents = "none";
  style.color = "transparent";
  style.background = "transparent";
  style.borderStyle = "solid";
  style.borderColor = "transparent";
  for (const property of MIRRORED_PROPERTIES) {
    style.setProperty(property, computed.getPropertyValue(property));
  }
  if (box instanceof HTMLInputElement) {
    // A single-line box never wraps; its words scroll sideways instead.
    style.whiteSpace = "pre";
    style.overflowWrap = "normal";
    style.wordBreak = "normal";
  }
  const text = box.value;
  const runs = (
    marks ?? [{ start: box.selectionStart ?? 0, end: box.selectionEnd ?? 0 }]
  )
    .map((mark) => ({
      ...mark,
      start: Math.max(0, Math.min(text.length, mark.start)),
      end: Math.max(0, Math.min(text.length, mark.end)),
    }))
    .filter((mark) => mark.end > mark.start)
    .sort((a, b) => a.start - b.start);
  let at = 0;
  for (const mark of runs) {
    if (mark.start < at) continue;
    ghost.append(root.createTextNode(text.slice(at, mark.start)));
    const range = root.createElement("span");
    range.setAttribute("data-office-inactive-selection-range", "");
    for (const [name, value] of Object.entries(mark.attributes ?? {})) range.setAttribute(name, value);
    if (mark.background) {
      // A highlight colour (a find hit) tints the words under it rather
      // than covering them: the mirror lies over the box.
      range.style.background = mark.background;
      range.style.mixBlendMode = "multiply";
    } else {
      // The system selection colour, strong enough to read on any shape
      // fill, translucent so the box's own words show through it.
      range.style.background = "rgba(0, 120, 215, 0.45)";
      range.style.background = "color-mix(in srgb, Highlight 70%, transparent)";
    }
    range.style.borderRadius = "1px";
    range.textContent = text.slice(mark.start, mark.end);
    ghost.append(range);
    at = mark.end;
  }
  ghost.append(root.createTextNode(text.slice(at)));
  if (beside) box.insertAdjacentElement("afterend", ghost);
  else root.body.append(ghost);
  ghost.scrollTop = box.scrollTop;
  ghost.scrollLeft = box.scrollLeft;
  return ghost;
}

/**
 * Paint runs of a plain text box's characters (a find hit, a mirrored
 * range) on its words, as the kept selection is painted: a mirror of the box
 * beside it, in its own place, size and transform, so the runs turn and
 * scale with the words, redrawn whenever the box's layout or text changes.
 * Nothing measures the words apart from the box (video audit H01). Returns
 * the remover, or null when the box is not drawn.
 */
export function officePaintTextControlMarks(
  box: HTMLElement,
  marks: readonly OfficeTextControlMark[],
): (() => void) | null {
  if (!isTextBox(box) || !boxShown(box) || marks.length === 0) return null;
  const root = box.ownerDocument;
  const view = root.defaultView;
  const marksKey = marks.map((mark) => `${mark.start}-${mark.end}-${mark.background ?? ""}`).join(",");
  const keyOf = () => `${ghostKey(box)}#${marksKey}`;
  let ghost = paintGhost(box, root, marks);
  if (!ghost) return null;
  let drawnKey = keyOf();
  let frame: number | null = null;
  const follow = () => {
    frame = null;
    if (!box.isConnected || !boxShown(box)) {
      ghost?.remove();
      ghost = null;
    } else {
      const key = keyOf();
      if (key !== drawnKey || !ghost) {
        ghost?.remove();
        ghost = paintGhost(box, root, marks);
        drawnKey = key;
      }
    }
    if (box.isConnected) {
      frame = typeof view?.requestAnimationFrame === "function" ? view.requestAnimationFrame(follow) : null;
    }
  };
  frame = typeof view?.requestAnimationFrame === "function" ? view.requestAnimationFrame(follow) : null;
  return () => {
    if (frame !== null && typeof view?.cancelAnimationFrame === "function") view.cancelAnimationFrame(frame);
    frame = null;
    ghost?.remove();
    ghost = null;
  };
}

/**
 * Install the rule on a document once; returns the uninstaller. Every text
 * box on the page is covered, so a host adds no markup.
 */
export function installOfficeInactiveSelection(root: Document = document): () => void {
  trackPresses(root);
  const view = root.defaultView;
  let ghost: HTMLElement | null = null;
  let ghosted: TextBox | null = null;
  let drawnKey = "";
  let frame: number | null = null;

  const clear = () => {
    if (frame !== null && typeof view?.cancelAnimationFrame === "function") view.cancelAnimationFrame(frame);
    frame = null;
    ghost?.remove();
    ghost = null;
    ghosted = null;
    drawnKey = "";
  };
  // While a range is kept, follow the box: redraw when its words, range or
  // place change (a ribbon command rewrote them, the page scrolled), drop the
  // drawing the moment the host closes the box (the deck unmounts its shape
  // editor on blur, Docs hides its member editor).
  const follow = () => {
    frame = null;
    const box = ghosted;
    if (!box) return;
    if (!boxShown(box) || !hasRange(box) || root.activeElement === box) {
      clear();
      return;
    }
    const key = ghostKey(box);
    if (key !== drawnKey) {
      ghost?.remove();
      ghost = paintGhost(box, root);
      drawnKey = key;
      if (!ghost) {
        clear();
        return;
      }
    }
    frame = typeof view?.requestAnimationFrame === "function" ? view.requestAnimationFrame(follow) : null;
  };
  const keep = (box: TextBox) => {
    clear();
    if (!boxShown(box) || !hasRange(box)) return;
    ghosted = box;
    follow();
  };
  const drop = (box: TextBox) => {
    if (ghosted === box) clear();
    if (!box.isConnected || !hasRange(box)) return;
    const end = box.selectionEnd ?? 0;
    try {
      box.setSelectionRange(end, end);
    } catch {
      // A box whose type has no selection API: nothing to collapse.
    }
    box.dispatchEvent(
      new CustomEvent(OFFICE_INACTIVE_SELECTION_COLLAPSE_EVENT, { bubbles: true }),
    );
  };

  const onFocusOut = (event: FocusEvent) => {
    const box = event.target;
    if (!isTextBox(box) || !hasRange(box)) return;
    if (box.hasAttribute(OFFICE_INACTIVE_SELECTION_SELF_PAINTED_ATTRIBUTE)) return;
    const related = event.relatedTarget;
    const intoChrome = officeBlurWentIntoChrome(related);
    // Decide once the focus (and the host's own blur handling, which may close
    // the box) has settled.
    view?.setTimeout(() => {
      if (root.activeElement === box) return;
      // The keyboard went to another text box: a jump from the chrome (the
      // Watch Window puts it in the target cell's entry) leaves the range,
      // though the press that started it was in the chrome (S-141).
      if (focusMovedToAnotherTextEntry(root.activeElement)) {
        drop(box);
        return;
      }
      if (intoChrome || !root.hasFocus() || inChrome(root.activeElement)) {
        // Into the ribbon, a menu, a pane -- or to another application:
        // the range is kept, and shown.
        keep(box);
        return;
      }
      // Into anything else: the user left the selection behind.
      drop(box);
    }, 0);
  };
  const onFocusIn = (event: FocusEvent) => {
    if (!ghosted) return;
    // Back in the box: the live selection draws itself again.
    if (event.target === ghosted) {
      clear();
      return;
    }
    // Into another text box or anywhere that is not chrome: the kept range
    // is left behind -- unless the move followed a press in the chrome (the
    // deck gives its slide canvas the focus when a ribbon tab is pressed;
    // PowerPoint keeps the text selection across a tab switch).
    if (focusMovedToAnotherTextEntry(event.target)) {
      drop(ghosted);
      return;
    }
    if (!inChrome(event.target) && !inChrome(officeRecentPressTarget())) drop(ghosted);
  };
  const onPress = (event: Event) => {
    if (!ghosted) return;
    const target = event.target;
    if (target === ghosted) {
      // Pressing the box itself takes the focus back; the live selection wins.
      clear();
      return;
    }
    // A press outside the chrome (the canvas, another object, the page)
    // leaves the selection, even where it gives nothing the focus.
    if (!inChrome(target)) drop(ghosted);
  };

  root.addEventListener("focusout", onFocusOut, true);
  root.addEventListener("focusin", onFocusIn, true);
  root.addEventListener("pointerdown", onPress, true);
  return () => {
    root.removeEventListener("focusout", onFocusOut, true);
    root.removeEventListener("focusin", onFocusIn, true);
    root.removeEventListener("pointerdown", onPress, true);
    clear();
  };
}

/**
 * Keep an in-place text editor open while the editor chrome has the pointer.
 *
 * An editor that closes itself on blur (a worksheet shape's words) closed the
 * moment a ribbon tab was pressed: the highlighted words vanished while the
 * shape stayed selected, which read as a selection that was cancelled yet
 * still in force. Called from the box's blur, this answers whether the blur
 * went into the chrome; if it did, the caller leaves the box open and
 * `onLeave` runs once the user really leaves -- a press outside both the box
 * and the chrome, or Escape. Returning to the box simply resumes the edit (its
 * next blur asks again).
 */
export function holdOfficeTextBoxForChrome<Box extends HTMLElement>(
  box: Box,
  related: EventTarget | null,
  onLeave: (box: Box) => void,
  options: {
    /**
     * Narrows what in the chrome keeps the box: a press on chrome this
     * rejects leaves the box first, as a press outside does (a cell entry is
     * held only for the character-formatting controls).
     */
    keepFor?: (target: EventTarget | null) => boolean;
  } = {},
): boolean {
  if (!officeBlurWentIntoChrome(related)) return false;
  const root = box.ownerDocument;
  // A rich (contenteditable) box has no selection of its own to keep: its
  // range is the document's, and the host may move that elsewhere while the
  // box is held (Docs returns the focus to the page after a ribbon command).
  // The range it had is painted as a highlight for as long as it is held.
  const view = root.defaultView;
  const selection = root.getSelection();
  const owner = Symbol("office-held-range");
  let painted = false;
  let watcher: MutationObserver | null = null;
  if (
    view && officeTextRangeHighlightsSupported() &&
    !(box instanceof HTMLTextAreaElement) && !(box instanceof HTMLInputElement) &&
    selection && selection.rangeCount > 0 && !selection.isCollapsed &&
    box.contains(selection.anchorNode) && box.contains(selection.focusNode)
  ) {
    // Kept as character offsets, not as the Range itself: a host that
    // rewrites the box's text on blur (the PDF text box commits it) replaces
    // the nodes a Range points into, and the band would collapse to nothing.
    const kept = selection.getRangeAt(0);
    const start = textOffsetIn(box, kept.startContainer, kept.startOffset);
    const end = textOffsetIn(box, kept.endContainer, kept.endOffset);
    const paint = () => {
      const range = rangeAtTextOffsets(box, start, end);
      setOfficeTextRangeHighlight(OFFICE_HELD_RANGE_PAINT, owner, range ? [toStaticRange(range)] : []);
    };
    paint();
    painted = true;
    if (typeof view.MutationObserver === "function") {
      watcher = new view.MutationObserver(paint);
      watcher.observe(box, { childList: true, subtree: true, characterData: true });
    }
  }
  const release = () => {
    root.removeEventListener("pointerdown", onPress, true);
    root.removeEventListener("keydown", onKey, true);
    box.removeEventListener("focusin", release);
    watcher?.disconnect();
    if (painted) clearOfficeTextRangeHighlight(OFFICE_HELD_RANGE_PAINT, owner);
  };
  const leave = () => {
    release();
    if (box.isConnected) onLeave(box);
  };
  const onPress = (event: Event) => {
    const target = event.target;
    if (!box.isConnected) {
      release();
      return;
    }
    if (target instanceof Node && box.contains(target)) return;
    // Before the press does its own work (selects a cell, another object):
    // the host commits the words to the object that is still selected.
    if (!inChrome(target) || (options.keepFor && !options.keepFor(target))) leave();
  };
  const onKey = (event: KeyboardEvent) => {
    if (!box.isConnected) release();
    else if (event.key === "Escape" && !inChrome(event.target)) leave();
  };
  root.addEventListener("pointerdown", onPress, true);
  root.addEventListener("keydown", onKey, true);
  // `focusin`, so a box that is a frame around its field (the document's
  // shape editor) resumes too.
  box.addEventListener("focusin", release);
  return true;
}

/** Characters of `box`'s text before (`node`, `offset`). */
function textOffsetIn(box: Node, node: Node, offset: number): number {
  const range = box.ownerDocument!.createRange();
  range.selectNodeContents(box);
  try {
    range.setEnd(node, offset);
  } catch {
    return 0;
  }
  return range.toString().length;
}

/** A Range over characters [start, end) of `box`'s text, or null when it has fewer. */
function rangeAtTextOffsets(box: Node, start: number, end: number): Range | null {
  const doc = box.ownerDocument!;
  const walker = doc.createTreeWalker(box, NodeFilter.SHOW_TEXT);
  const range = doc.createRange();
  let seen = 0;
  let started = false;
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const length = node.textContent?.length ?? 0;
    if (!started && start <= seen + length) {
      range.setStart(node, start - seen);
      started = true;
    }
    if (started && end <= seen + length) {
      range.setEnd(node, end - seen);
      return range.collapsed ? null : range;
    }
    seen += length;
  }
  return null;
}

/**
 * The shared rule's answer for a box that paints its own selection (marked
 * with OFFICE_INACTIVE_SELECTION_SELF_PAINTED_ATTRIBUTE): `onChange(true)`
 * once the focus has left the box with characters selected for the editor
 * chrome (the ribbon, a font list, a dialog) or for another application --
 * the range is held and stays painted, as Word and WPS keep it; `false`
 * when the focus comes back to the box, or a press or the focus goes
 * anywhere that is not chrome. Returns the remover.
 */
export function watchOfficeInactiveSelection(
  box: HTMLTextAreaElement | HTMLInputElement,
  onChange: (held: boolean) => void,
): () => void {
  const root = box.ownerDocument;
  const view = root.defaultView;
  trackPresses(root);
  let held = false;
  let timer: number | null = null;
  const set = (next: boolean) => {
    if (next === held) return;
    held = next;
    onChange(next);
  };
  const onBlur = (event: FocusEvent) => {
    if (!hasRange(box)) return;
    const intoChrome = officeBlurWentIntoChrome(event.relatedTarget);
    if (timer !== null) view?.clearTimeout(timer);
    // Decided once the focus (and the host's own blur handling) has settled.
    timer = view?.setTimeout(() => {
      timer = null;
      if (root.activeElement === box) return;
      if (focusMovedToAnotherTextEntry(root.activeElement)) {
        set(false);
        return;
      }
      set(
        boxShown(box) && hasRange(box) &&
          (intoChrome || !root.hasFocus() || inChrome(root.activeElement)),
      );
    }, 0) ?? null;
  };
  const onFocus = () => set(false);
  const onFocusIn = (event: FocusEvent) => {
    if (!held || event.target === box) return;
    if (focusMovedToAnotherTextEntry(event.target)) {
      set(false);
      return;
    }
    if (!inChrome(event.target) && !inChrome(officeRecentPressTarget())) set(false);
  };
  const onPress = (event: Event) => {
    if (!held) return;
    if (event.target === box || !inChrome(event.target)) set(false);
  };
  // One element type for the listener overloads (a union of the two picks
  // the untyped one).
  const field: HTMLElement = box;
  field.addEventListener("blur", onBlur);
  field.addEventListener("focus", onFocus);
  root.addEventListener("focusin", onFocusIn, true);
  root.addEventListener("pointerdown", onPress, true);
  return () => {
    if (timer !== null) view?.clearTimeout(timer);
    field.removeEventListener("blur", onBlur);
    field.removeEventListener("focus", onFocus);
    root.removeEventListener("focusin", onFocusIn, true);
    root.removeEventListener("pointerdown", onPress, true);
  };
}

let installs = 0;
let uninstall: (() => void) | null = null;

const OFFICE_MIRRORED_RANGE_PAINT: OfficeTextRangeHighlightPaint = {
  name: "office-mirrored-range",
  css: "background-color: color-mix(in srgb, Highlight 70%, transparent);",
};

/**
 * Shows, on a box that MIRRORS another editor's text (a cell entry and the
 * formula bar are one entry in two places), the characters `start..end`
 * selected in that other editor -- as Excel and WPS show the formula bar's
 * selection in the cell and the cell's in the formula bar. The box keeps its
 * own focus state: a plain input or text area gets the selection ghost (its
 * own selection is set to the range, which does not take the focus), a rich
 * box a CSS highlight over the range. Returns the remover.
 */
export function officeShowMirroredSelection(
  box: HTMLElement,
  start: number,
  end: number,
): (() => void) | null {
  const root = box.ownerDocument;
  if (!box.isConnected || root.activeElement === box) return null;
  const from = Math.min(start, end);
  const to = Math.max(start, end);
  if (to <= from) return null;
  if (isTextBox(box)) {
    try {
      box.setSelectionRange(from, to);
    } catch {
      return null;
    }
    if (!boxShown(box) || !hasRange(box)) return null;
    // Drawn from the box's layout and redrawn whenever that layout changes
    // (a new size or font, a resized row, a scroll, a font that finished
    // loading), as the kept range is: a mirror drawn once stayed where the
    // words had been.
    const view = root.defaultView;
    let ghost = paintGhost(box, root);
    if (!ghost) return null;
    let drawnKey = ghostKey(box);
    let frame: number | null = null;
    const follow = () => {
      frame = null;
      if (!box.isConnected || !boxShown(box) || root.activeElement === box) {
        ghost?.remove();
        ghost = null;
        return;
      }
      const key = ghostKey(box);
      if (key !== drawnKey) {
        ghost?.remove();
        ghost = hasRange(box) ? paintGhost(box, root) : null;
        drawnKey = key;
        if (!ghost) return;
      }
      frame = typeof view?.requestAnimationFrame === "function" ? view.requestAnimationFrame(follow) : null;
    };
    frame = typeof view?.requestAnimationFrame === "function" ? view.requestAnimationFrame(follow) : null;
    return () => {
      if (frame !== null && typeof view?.cancelAnimationFrame === "function") view.cancelAnimationFrame(frame);
      frame = null;
      ghost?.remove();
      ghost = null;
    };
  }
  if (!officeTextRangeHighlightsSupported()) return null;
  // The words themselves are coloured (a CSS highlight over their own text),
  // so the band cannot leave them. Each mirror owns its range: two windows
  // showing a mirrored range no longer clear each other's.
  const owner = Symbol("office-mirrored-range");
  const bind = (): boolean => {
    const range = rangeAtTextOffsets(box, from, to);
    setOfficeTextRangeHighlight(
      OFFICE_MIRRORED_RANGE_PAINT,
      owner,
      range ? [toStaticRange(range)] : [],
    );
    return range !== null;
  };
  if (!bind()) return null;
  // A re-render that replaces the words' text nodes leaves a static range on
  // nodes that are gone: the range is found again in the new ones.
  const view = root.defaultView;
  const observer =
    view && typeof view.MutationObserver === "function" ? new view.MutationObserver(() => bind()) : null;
  observer?.observe(box, { subtree: true, childList: true, characterData: true });
  return () => {
    observer?.disconnect();
    clearOfficeTextRangeHighlight(OFFICE_MIRRORED_RANGE_PAINT, owner);
  };
}

/** Reference-counted install for the editor windows that share one page. */
export function retainOfficeInactiveSelection(): () => void {
  if (typeof document === "undefined") return () => {};
  installs += 1;
  if (installs === 1) uninstall = installOfficeInactiveSelection(document);
  let released = false;
  return () => {
    if (released) return;
    released = true;
    installs -= 1;
    if (installs === 0) {
      uninstall?.();
      uninstall = null;
    }
  };
}

// The press that moves the focus happens before any editor mounts the rule;
// track presses from the first import so the question can always be answered.
if (typeof document !== "undefined") trackPresses(document);
