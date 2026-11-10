/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * An equation in a slide's text as the editing surface sees it: one ATOM.
 *
 * A run that is an equation (`TextRun.math`) is drawn by KaTeX, so the DOM
 * inside its element is glyph markup and not the run's characters. The text
 * readers of a text box therefore never read inside such an element: they take
 * its characters from `data-presentation-math-text` (the run's `text`, what
 * the server's Office Math draws), count it as that many characters, and keep
 * the caret outside it -- before or after, never between its glyphs. Pure DOM
 * helpers, shared by the renderer (`PresentationMathRun`) and the selection
 * readers (`text-selection`).
 */

/** The attribute the text readers take an equation's characters from. */
export const PRESENTATION_MATH_TEXT_ATTRIBUTE = "data-presentation-math-text";
/** The selector of an equation atom in a text box's DOM. */
export const PRESENTATION_MATH_SELECTOR = "[data-presentation-math]";

function isElement(node: Node | null | undefined): node is Element {
  return !!node && node.nodeType === Node.ELEMENT_NODE;
}

/** The equation atom `node` is or lies inside, else null. */
export function presentationMathAtomOf(node: Node | null | undefined): HTMLElement | null {
  const element = isElement(node) ? node : node?.parentElement ?? null;
  return element?.closest<HTMLElement>(PRESENTATION_MATH_SELECTOR) ?? null;
}

/** Whether `element` is an equation atom. */
export function isPresentationMathAtom(node: Node | null | undefined): node is HTMLElement {
  return isElement(node) && node.matches(PRESENTATION_MATH_SELECTOR);
}

/** The characters an equation atom stands for in the block's text. */
export function presentationMathAtomText(atom: Element): string {
  return atom.getAttribute(PRESENTATION_MATH_TEXT_ATTRIBUTE) ?? "";
}

/**
 * A FIELD drawn as something other than its stored characters -- the slide
 * number, stored as the number the field had when it was typed and drawn
 * as the slide's current one -- is an atom to the text readers too: it
 * counts as its stored characters, so a selection read off the drawn "10"
 * lands on the stored "9" and not on the letter after it (PPT-B429).
 */
export const PRESENTATION_FIELD_TEXT_ATTRIBUTE = "data-presentation-field-text";
const PRESENTATION_TEXT_ATOM_SELECTOR = `${PRESENTATION_MATH_SELECTOR}, [${PRESENTATION_FIELD_TEXT_ATTRIBUTE}]`;

/** An equation or a field: what the text readers count by its stored characters. */
function isPresentationTextAtom(node: Node | null | undefined): node is HTMLElement {
  return isElement(node) && node.matches(PRESENTATION_TEXT_ATOM_SELECTOR);
}

function presentationTextAtomText(atom: Element): string {
  return (
    atom.getAttribute(PRESENTATION_MATH_TEXT_ATTRIBUTE) ??
    atom.getAttribute(PRESENTATION_FIELD_TEXT_ATTRIBUTE) ??
    ""
  );
}

/**
 * A paragraph's drawn MARKER (its bullet or number) inside an editing host
 * that types the words where they are drawn: it is the paragraph's, not its
 * characters, so the text readers count it as nothing and never put the
 * caret inside or before it.
 */
export const OFFICE_TEXT_MARKER_ATTRIBUTE = "data-office-text-marker";

function isOfficeTextMarker(node: Node | null | undefined): boolean {
  return isElement(node) && node.hasAttribute(OFFICE_TEXT_MARKER_ATTRIBUTE);
}

/**
 * The element holding the zero-width character a text box draws after a
 * final line break, so the browser lays the empty last line out
 * (`PresentationRunsText`). It is not text: the readers skip that character,
 * or typing on the new line saved it (acceptance 43b: "def\nzz\u200b").
 */
export const PRESENTATION_TRAILING_LINE_BOX_ATTRIBUTE = "data-trailing-line-box";
const TRAILING_LINE_BOX_CHARACTER = "\u200b";

function inTrailingLineBox(text: Text): boolean {
  return !!text.parentElement?.closest(`[${PRESENTATION_TRAILING_LINE_BOX_ATTRIBUTE}]`);
}

/** A text node's characters as the model counts them. */
function modelTextOf(text: Text): string {
  return inTrailingLineBox(text)
    ? text.data.split(TRAILING_LINE_BOX_CHARACTER).join("")
    : text.data;
}

/** The model characters before DOM offset `domOffset` of a text node. */
function modelLengthBefore(text: Text, domOffset: number): number {
  const before = text.data.slice(0, Math.max(0, Math.min(domOffset, text.data.length)));
  return inTrailingLineBox(text)
    ? before.split(TRAILING_LINE_BOX_CHARACTER).join("").length
    : before.length;
}

/** The DOM offset of model offset `modelOffset` inside a text node. */
function domOffsetOf(text: Text, modelOffset: number): number {
  if (!inTrailingLineBox(text)) return modelOffset;
  let seen = 0;
  for (let index = 0; index < text.data.length; index += 1) {
    if (seen === modelOffset) return index;
    if (text.data[index] !== TRAILING_LINE_BOX_CHARACTER) seen += 1;
  }
  return text.data.length;
}

/**
 * The characters of `root` as the text model counts them: text nodes as they
 * are (less a trailing line box's placeholder), an equation atom as its
 * characters, nothing else.
 */
export function presentationAtomAwareText(root: Node): string {
  if (root.nodeType === Node.TEXT_NODE) return modelTextOf(root as Text);
  if (!isElement(root) || isOfficeTextMarker(root)) return "";
  if (isPresentationTextAtom(root)) return presentationTextAtomText(root);
  let out = "";
  for (const child of Array.from(root.childNodes)) out += presentationAtomAwareText(child);
  return out;
}

/**
 * The model offset of the DOM point (`node`, `nodeOffset`) inside `root`, in
 * the characters {@link presentationAtomAwareText} counts. A point inside an
 * equation atom is the atom's end when it lies in the atom's second half,
 * else its start. A point outside `root` reads as `root`'s whole length.
 */
export function presentationAtomAwareOffset(root: Element, node: Node, nodeOffset: number): number {
  let total = 0;
  const walk = (current: Node): boolean => {
    if (current === node) {
      if (isOfficeTextMarker(current)) return true;
      if (current.nodeType === Node.TEXT_NODE) {
        total += modelLengthBefore(current as Text, nodeOffset);
        return true;
      }
      if (isPresentationTextAtom(current)) {
        total += nodeOffset > 0 ? presentationTextAtomText(current).length : 0;
        return true;
      }
      if (isElement(current)) {
        // An element point: the children before `nodeOffset` are behind it.
        for (const child of Array.from(current.childNodes).slice(0, Math.max(0, nodeOffset))) {
          total += presentationAtomAwareText(child).length;
        }
      }
      return true;
    }
    if (current.nodeType === Node.TEXT_NODE) {
      total += modelTextOf(current as Text).length;
      return false;
    }
    if (!isElement(current)) return false;
    // A marker: no characters; a point in it is the paragraph's start.
    if (isOfficeTextMarker(current)) return current.contains(node);
    if (isPresentationTextAtom(current)) {
      const length = presentationTextAtomText(current).length;
      if (current.contains(node)) {
        total += atomHalf(current, node, nodeOffset) ? length : 0;
        return true;
      }
      total += length;
      return false;
    }
    for (const child of Array.from(current.childNodes)) if (walk(child)) return true;
    return false;
  };
  walk(root);
  return total;
}

/** Whether a point inside an equation atom lies in its second half. */
function atomHalf(atom: Element, node: Node, nodeOffset: number): boolean {
  const inner = atom.textContent ?? "";
  if (!inner.length) return nodeOffset > 0;
  try {
    const probe = atom.ownerDocument.createRange();
    probe.selectNodeContents(atom);
    probe.setEnd(node, nodeOffset);
    return probe.toString().length * 2 >= inner.length;
  } catch {
    return nodeOffset > 0;
  }
}

/** A DOM point the browser can put a caret on. */
export type PresentationDomPoint = Readonly<{ node: Node; offset: number }>;














export function presentationAtomAwareDomPoint(root: Element, offset: number): PresentationDomPoint {
  let remaining = Math.max(0, Math.trunc(offset));
  let last: PresentationDomPoint | null = null;
  const visit = (element: Element): PresentationDomPoint | null => {
    const children = Array.from(element.childNodes);
    for (let index = 0; index < children.length; index += 1) {
      const child = children[index]!;
      if (child.nodeType === Node.TEXT_NODE) {
        const text = child as Text;
        const model = modelTextOf(text);
        const length = model.length;
        if (remaining < length || (remaining === length && !model.endsWith("\n"))) {
          return { node: text, offset: domOffsetOf(text, remaining) };
        }
        remaining -= length;
        last = { node: text, offset: domOffsetOf(text, length) };
        continue;
      }
      if (!isElement(child) || isOfficeTextMarker(child)) continue;
      // A line break that holds a line with no text (an empty paragraph, a
      // line a final break opens): the start of that line is before it.
      if (child.nodeName === "BR" && remaining === 0) return { node: element, offset: index };
      if (isPresentationTextAtom(child)) {
        if (remaining <= 0) return { node: element, offset: index };
        const length = presentationTextAtomText(child).length;
        remaining = Math.max(0, remaining - length);
        last = { node: element, offset: index + 1 };
        if (remaining === 0) {
          // Right after the equation: inside its caret anchor, past the
          // zero-width character, where the browser keeps a caret.
          const anchor = children[index + 1];
          const anchorText = isElement(anchor) && anchor.hasAttribute(PRESENTATION_TRAILING_LINE_BOX_ATTRIBUTE)
            ? anchor.firstChild
            : null;
          if (anchorText && anchorText.nodeType === Node.TEXT_NODE) {
            return { node: anchorText, offset: (anchorText as Text).length };
          }
          return last;
        }
        continue;
      }
      const found = visit(child);
      if (found) return found;
    }
    return null;
  };
  return visit(root) ?? last ?? { node: root, offset: 0 };
}
