/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Text-selection helpers of the presentation editor: what the Ribbon reads for
 * a caret or a selection, the Format Painter's paragraph ranges, proofing and
 * spellcheck state for a text block, and the mapping between plain-text offsets
 * and the DOM of the non-editing text preview.
 *
 * Moved verbatim out of `components/notebook/contents/PresentationWindow.tsx`
 * (structural split, no behaviour change).
 */

import { officeTextOffsetAtPoint } from "../text-engine/office-text-hit";
import { officeCaretBox } from "../office-painted-caret";
import {
  presentationAtomAwareDomPoint,
  presentationAtomAwareOffset,
  presentationAtomAwareText,
  type PresentationDomPoint,
} from "./math-atom";
import {
  type ParagraphAlign,
  clampLevel,
  normalizeBullet,
  textFromParagraphs,
} from "../presentation-paragraphs";
import {
  type TextRun,
  mergeRuns,
  normalizeRunStyle,
  normalizeRuns,
  paragraphRunSelectionRanges,
  runFieldAtSelection,
  runProofingAtSelection,
  runsFromText,
  textFromRuns,
} from "../presentation-point-runs";
import type {
  FormatPainterParagraphStylePatch,
  FormatPainterRunRange,
} from "../presentation-format-painter";
import {
  applyParagraphFormatSelection,
  type LiveTextSelection,
  normalizedSpacingFallbackBlocked,
  normalizeParagraphsWithRuns,
  ownsParagraphRichLayout,
  paragraphFormatValue,
  paragraphPersistenceEqual,
  paragraphSelectionSpan,
  paragraphsWithRunsFromTextBlock,
  PRESENTATION_PARAGRAPH_FORMAT_FIELDS,
  PRESENTATION_PARAGRAPH_SPACING_FIELDS,
  type PresentationParagraphFormatPatch,
  selectedTextRange,
  type SlideParagraphWithRuns,
  withPersistedParagraphs,
} from "../presentation-text-buttons";
import type {
  PresentationTextBlock,
  PresentationTextStyle,
} from "../../components/console/editor/native-document/documentCommands";

/**
 * What the Ribbon READS for its state: {@link selectedTextRange}, except that a
 * caret owning no word reads the character beside it on its own line (before
 * it, or after it at a line start), the formatting typing there continues.
 */
export function selectedTextReadRange(
  blockId: string,
  text: string,
  selection: { blockId: string; start: number; end: number } | null,
): readonly [number, number] {
  const [from, to] = selectedTextRange(blockId, text, selection);
  if (to > from) return [from, to];
  if (from > 0 && text[from - 1] !== "\n") return [from - 1, from];
  if (from < text.length && text[from] !== "\n") return [from, from + 1];
  return [from, to];
}

export type EffectiveFontFamilyState =
  | Readonly<{ kind: "uniform"; family: string }>
  | Readonly<{ kind: "mixed" }>;

/**
 * Resolve run inheritance before comparing the selected faces. An absent run
 * override means "inherit", not "unknown"; mapping that absence to the live
 * block/theme family lets the Ribbon distinguish a genuinely mixed selection
 * from uniformly inherited text.
 */
export function effectiveFontFamilyForRunRange(
  runs: readonly TextRun[],
  start: number,
  end: number,
  inheritedFamily: string,
): EffectiveFontFamilyState {
  const base = mergeRuns(runs);
  if (textFromRuns(base).length === 0) {
    return { kind: "uniform", family: inheritedFamily };
  }
  const resolvedRuns = base.map((run) => ({
    text: run.text,
    style: {
      ...(normalizeRunStyle(run.style) ?? {}),
      fontFamily:
        normalizeRunStyle(run.style)?.fontFamily ?? inheritedFamily,
    },
  }));
  const family = runFieldAtSelection(
    resolvedRuns,
    start,
    end,
    "fontFamily",
  );
  return typeof family === "string"
    ? { kind: "uniform", family }
    : { kind: "mixed" };
}

/**
 * Add paragraph formatting authority only when the selected text includes the
 * complete paragraph plus its paragraph mark. Partial text keeps the existing
 * character-only Format Painter contract.
 */
export function formatPainterSelectionRanges(
  paragraphs: readonly SlideParagraphWithRuns[],
  from: number,
  to: number,
  fallback: PresentationTextStyle | undefined,
): FormatPainterRunRange[] {
  const starts: number[] = [];
  let cursor = 0;
  paragraphs.forEach((paragraph, index) => {
    starts.push(cursor);
    cursor += paragraph.text.length + (index < paragraphs.length - 1 ? 1 : 0);
  });
  return paragraphRunSelectionRanges(paragraphs, from, to).map((range) => {
    const paragraph = paragraphs[range.paragraphIndex];
    const start = starts[range.paragraphIndex] ?? 0;
    const textEnd = start + (paragraph?.text.length ?? 0);
    const markEnd =
      textEnd + (range.paragraphIndex < paragraphs.length - 1 ? 1 : 0);
    if (!paragraph || from > start || to < markEnd) return range;

    const blocked = new Set(
      normalizedSpacingFallbackBlocked(paragraph.spacingFallbackBlocked) ?? [],
    );
    const style: FormatPainterParagraphStylePatch = {
      level: clampLevel(paragraph.level),
      bullet: normalizeBullet(paragraph.bullet),
      align: paragraphFormatValue(
        paragraph,
        "align",
        fallback,
      ) as ParagraphAlign,
    };
    for (const field of PRESENTATION_PARAGRAPH_SPACING_FIELDS) {
      if (blocked.has(field)) continue;
      const value = paragraphFormatValue(paragraph, field, fallback) as
        | number
        | undefined;
      style[field] = value ?? null;
    }
    return { ...range, paragraph: style };
  });
}

/** Apply one authenticated Format Painter paragraph patch to the live span. */
export function applyFormatPainterParagraphSelection(
  block: PresentationTextBlock,
  selection: LiveTextSelection,
  patch: FormatPainterParagraphStylePatch,
): PresentationTextBlock | null {
  const paragraphPatch: PresentationParagraphFormatPatch = {};
  for (const field of PRESENTATION_PARAGRAPH_FORMAT_FIELDS) {
    if (!Object.prototype.hasOwnProperty.call(patch, field)) continue;
    if (field === "align") paragraphPatch.align = patch.align;
    if (field === "lineSpacing") {
      paragraphPatch.lineSpacing = patch.lineSpacing ?? undefined;
    }
    if (field === "spaceBefore") {
      paragraphPatch.spaceBefore = patch.spaceBefore ?? undefined;
    }
    if (field === "spaceAfter") {
      paragraphPatch.spaceAfter = patch.spaceAfter ?? undefined;
    }
  }

  const formatted =
    applyParagraphFormatSelection(block, selection, paragraphPatch) ?? block;
  const ownsLevel = Object.prototype.hasOwnProperty.call(patch, "level");
  const ownsBullet = Object.prototype.hasOwnProperty.call(patch, "bullet");
  if (!ownsLevel && !ownsBullet) {
    return formatted === block ? null : formatted;
  }

  const paragraphs = paragraphsWithRunsFromTextBlock(formatted);
  if (paragraphs.length === 0) return formatted === block ? null : formatted;
  const [first, last] = paragraphSelectionSpan(
    formatted.id,
    paragraphs,
    selection,
  );
  let changed = false;
  const nextParagraphs = paragraphs.map((paragraph, index) => {
    if (index < first || index > last) return paragraph;
    const level = ownsLevel ? clampLevel(patch.level) : paragraph.level;
    const bullet = ownsBullet
      ? normalizeBullet(patch.bullet)
      : paragraph.bullet;
    if (level === paragraph.level && bullet === paragraph.bullet) {
      return paragraph;
    }
    changed = true;
    return { ...paragraph, level, bullet };
  });
  const next = changed
    ? withPersistedParagraphs(formatted, nextParagraphs)
    : formatted;
  return paragraphPersistenceEqual(next, block) ? null : next;
}

export function proofingStateForTextBlockSelection(
  textBlock: PresentationTextBlock,
  selection: LiveTextSelection,
): {
  readonly proofingLanguage?: string;
  readonly noProof: boolean;
  readonly mixed: boolean;
} {
  const paragraphValues = normalizeParagraphsWithRuns(textBlock.paragraphs);
  const states = paragraphValues && ownsParagraphRichLayout(paragraphValues)
    ? paragraphRunSelectionRanges(
        paragraphValues,
        ...selectedTextReadRange(
          textBlock.id,
          textFromParagraphs(paragraphValues),
          selection,
        ),
      ).map((range) =>
        runProofingAtSelection(range.runs, range.from, range.to),
      )
    : [
        runProofingAtSelection(
          normalizeRuns(textBlock.runs) ?? runsFromText(textBlock.text),
          ...selectedTextReadRange(
            textBlock.id,
            textBlock.text,
            selection,
          ),
        ),
      ];
  if (states.length === 0 || states.some((state) => state.mixed)) {
    return { noProof: false, mixed: true };
  }
  const keys = new Set(
    states.map(
      (state) =>
        `${state.value?.proofingLanguage ?? ""}\u0000${state.value?.noProof === true}`,
    ),
  );
  if (keys.size !== 1) return { noProof: false, mixed: true };
  const value = states[0]?.value;
  return {
    ...(value?.proofingLanguage
      ? { proofingLanguage: value.proofingLanguage }
      : {}),
    noProof: value?.noProof === true,
    mixed: false,
  };
}

export function browserSpellcheckForTextBlock(
  textBlock: PresentationTextBlock,
): { readonly spellCheck: boolean; readonly lang?: string } {
  const paragraphs = normalizeParagraphsWithRuns(textBlock.paragraphs);
  const runs =
    paragraphs && ownsParagraphRichLayout(paragraphs)
      ? paragraphs.flatMap(
          (paragraph) =>
            normalizeRuns(paragraph.runs) ?? runsFromText(paragraph.text),
        )
      : normalizeRuns(textBlock.runs) ?? runsFromText(textBlock.text);
  // A textarea/contenteditable cannot suppress spelling for only part of one
  // mixed run graph. Honour any explicit DrawingML noProof range by failing the
  // whole block closed rather than underlining text PowerPoint excluded.
  if (runs.some((run) => run.style?.noProof === true)) {
    return { spellCheck: false };
  }
  const languages = new Set(
    runs.map((run) => run.style?.proofingLanguage ?? ""),
  );
  const language = languages.size === 1 ? [...languages][0] : "";
  return {
    spellCheck: true,
    ...(language ? { lang: language } : {}),
  };
}

type PresentationPreviewSelection = Readonly<{
  start: number;
  end: number;
}>;

/**
 * The characters a paragraph's text root draws, as the model counts them: an
 * equation atom counts as the characters it stands for, never as the KaTeX
 * markup inside it (`lib/presentation/math-atom`).
 */
function liveTextOf(textRoot: Element): string {
  return presentationAtomAwareText(textRoot);
}

/**
 * Whether a paragraph row reading `text` is an empty row's sentinel. At rest
 * an empty row keeps its height with a space (`PresentationParagraphList`),
 * which is no text. A body being typed in draws its empty rows as a line
 * break instead (`emptyRowAsBreak`): a space typed beside the drawn one was
 * kept with it ("abc\nx "), and a space typed alone in a new paragraph read
 * as the sentinel -- dropped, with the caret put back before it. There a
 * space is always one the person typed.
 */
function isEmptyRowSentinel(textRoot: Element, text: string): boolean {
  return (
    text === " " &&
    !textRoot.closest('[contenteditable="true"], [contenteditable=""], [contenteditable="plaintext-only"]')
  );
}

/** A paragraph row's words as the model counts them (its sentinel is none). */
function paragraphRowText(textRoot: Element | null): string {
  if (!textRoot) return "";
  const text = liveTextOf(textRoot);
  return isEmptyRowSentinel(textRoot, text) ? "" : text;
}

function presentationPreviewPointOffset(
  root: HTMLElement,
  node: Node,
  nodeOffset: number,
  paragraphs: readonly SlideParagraphWithRuns[] | undefined,
): number {
  const element =
    node.nodeType === Node.ELEMENT_NODE
      ? (node as Element)
      : node.parentElement;
  const paragraphElement = element?.closest<HTMLElement>(
    "[data-presentation-paragraph-index]",
  );
  // The paragraph rows (a list's bullets and numbers are drawn, not text):
  // counted by their words, with or without the model's paragraphs at hand.
  // Without them the count ran through the drawn marks, and the caret put
  // back after a keystroke stood two places on (a focus handover after the
  // body re-rendered), so typed words came out scrambled.
  if (paragraphElement && root.contains(paragraphElement)) {
    const paragraphIndex = Number(
      paragraphElement.dataset.presentationParagraphIndex,
    );
    const paragraph = paragraphs ? paragraphs[paragraphIndex] : true;
    if (paragraph) {
      const liveParagraphs = Array.from(
        root.querySelectorAll<HTMLElement>(
          "[data-presentation-paragraph-index]",
        ),
      );
      const liveParagraphText = (item: HTMLElement): string => {
        const textRoot = item.querySelector<HTMLElement>(
          '[data-presentation-paragraph-text="true"]',
        );
        const text = textRoot ? liveTextOf(textRoot) : "";
        // An empty row at rest draws a space to keep its height (the plain
        // text, `presentationPreviewPlainText`, reads it as empty).
        return textRoot &&
          isEmptyRowSentinel(textRoot, text) &&
          (!paragraphs ||
            paragraphs[Number(item.dataset.presentationParagraphIndex)]?.text ===
              "")
          ? ""
          : text;
      };
      const paragraphBase = liveParagraphs
        .slice(0, paragraphIndex)
        .reduce((total, item) => total + liveParagraphText(item).length + 1, 0);
      const textRoot = paragraphElement.querySelector<HTMLElement>(
        '[data-presentation-paragraph-text="true"]',
      );
      if (textRoot && (node === textRoot || textRoot.contains(node))) {
        const liveText = liveParagraphText(paragraphElement);
        return (
          paragraphBase +
          Math.min(
            liveText.length,
            presentationAtomAwareOffset(textRoot, node, nodeOffset),
          )
        );
      }
      // A drag that begins on the marker/gutter belongs to the start of that
      // paragraph. An element-boundary endpoint after the text belongs to its end.
      return (
        paragraphBase +
        (node === paragraphElement && nodeOffset > 0
          ? liveParagraphText(paragraphElement).length
          : 0)
      );
    }
  }

  // Flat run previews expose the exact block text in DOM order (an equation
  // atom as its characters), so the prefix length is the precise model offset.
  return presentationAtomAwareOffset(root, node, nodeOffset);
}

/**
 * The text offset a click at (`clientX`, `clientY`) lands on in a drawn
 * preview, or null when the point misses its text. A preview drawn over its
 * editing field takes no pointer; it answers the hit test for this call
 * alone, with the elements in `transparent` (the field) let through.
 */
export function presentationPreviewOffsetAtPoint(
  root: HTMLElement,
  clientX: number,
  clientY: number,
  textLength: number,
  paragraphs?: readonly SlideParagraphWithRuns[],
  transparent: readonly HTMLElement[] = [],
): number | null {
  // Words the text engine drew (SVG): their character from its own lines.
  const engineOffset = officeTextOffsetAtPoint(root, clientX, clientY);
  if (engineOffset !== null) return Math.max(0, Math.min(textLength, engineOffset));
  const doc = root.ownerDocument as Document & {
    caretRangeFromPoint?: (x: number, y: number) => Range | null;
    caretPositionFromPoint?: (
      x: number,
      y: number,
    ) => { offsetNode: Node; offset: number } | null;
  };
  const restore: Array<[HTMLElement, string]> = [];
  const setPointer = (element: HTMLElement, value: string) => {
    // MathML has no inline style in every engine (jsdom).
    if (!element.style) return;
    restore.push([element, element.style.pointerEvents]);
    element.style.pointerEvents = value;
  };
  setPointer(root, "auto");
  root
    .querySelectorAll<HTMLElement>("*")
    .forEach((element) => setPointer(element, "auto"));
  transparent.forEach((element) => setPointer(element, "none"));
  let node: Node | null = null;
  let nodeOffset = 0;
  try {
    const range = doc.caretRangeFromPoint?.(clientX, clientY);
    if (range) {
      node = range.startContainer;
      nodeOffset = range.startOffset;
    } else {
      const position = doc.caretPositionFromPoint?.(clientX, clientY);
      if (position) {
        node = position.offsetNode;
        nodeOffset = position.offset;
      }
    }
  } finally {
    for (const [element, value] of restore.reverse()) {
      element.style.pointerEvents = value;
    }
  }
  if (!node || !root.contains(node)) return null;
  return Math.max(
    0,
    Math.min(
      textLength,
      Math.trunc(
        presentationPreviewPointOffset(root, node, nodeOffset, paragraphs),
      ),
    ),
  );
}












export function presentationDoubleClickWord(text: string, offset: number): [number, number] {
  const at = Math.max(0, Math.min(text.length, Math.trunc(offset)));
  const lineStart = text.lastIndexOf("\n", at - 1) + 1;
  const breakAt = text.indexOf("\n", at);
  const lineEnd = breakAt < 0 ? text.length : breakAt;
  const probe = at < lineEnd ? at : at - 1;
  if (probe < lineStart) return [at, at];
  const [start, end] = wordAround(text, probe);
  return [Math.max(lineStart, start), Math.min(lineEnd, Math.max(end, probe + 1))];
}

/** The word (or run of one script's letters) around `offset` in `text`. */
function wordAround(text: string, offset: number): [number, number] {
  const Segmenter = (Intl as typeof Intl & {
    Segmenter?: new (
      locale?: string,
      options?: { granularity: "word" },
    ) => { segment(input: string): Iterable<{ segment: string; index: number; isWordLike?: boolean }> };
  }).Segmenter;
  if (Segmenter) {
    for (const part of new Segmenter(undefined, { granularity: "word" }).segment(text)) {
      if (offset >= part.index && offset < part.index + part.segment.length) {
        return [part.index, part.index + part.segment.length];
      }
    }
    return [offset, offset];
  }
  let start = offset;
  let end = offset;
  while (start > 0 && /\w/u.test(text[start - 1])) start -= 1;
  while (end < text.length && /\w/u.test(text[end])) end += 1;
  return [start, end];
}















export function followPreviewTextPress(
  field: HTMLTextAreaElement,
  preview: HTMLElement | (() => HTMLElement | null),
  event: PointerEvent,
  settled: (start: number, end: number) => void,
  pressed?: (start: number, end: number) => void,
): boolean {
  if (event.button !== 0 || event.pointerType === "touch" || !field.value) return false;
  const length = field.value.length;
  // The preview is read again at every step: selecting the box remounts
  // the host that draws its words, mid-drag.
  const offsetAt = (x: number, y: number) => {
    const root = typeof preview === "function" ? preview() : preview;
    return root
      ? presentationPreviewOffsetAtPoint(root, x, y, length, undefined, [field])
      : null;
  };
  const hit = offsetAt(event.clientX, event.clientY);
  if (hit === null) return false;
  // The field's own hit test would answer the press from its plain layout.
  event.preventDefault();
  const doc = field.ownerDocument;
  const extending = event.shiftKey && doc.activeElement === field;
  let anchor = extending
    ? field.selectionDirection === "backward"
      ? field.selectionEnd ?? hit
      : field.selectionStart ?? hit
    : hit;
  field.focus({ preventScroll: true });
  if (event.detail >= 2) {
    const text = field.value;
    const [start, end] =
      event.detail === 2
        ? presentationDoubleClickWord(text, hit)
        : presentationParagraphRange(text, hit);
    field.setSelectionRange(start, end);
    settled(start, end);
    return true;
  }
  let focus = hit;
  const select = () =>
    field.setSelectionRange(
      Math.min(anchor, focus),
      Math.max(anchor, focus),
      focus < anchor ? "backward" : "forward",
    );
  select();
  pressed?.(Math.min(anchor, focus), Math.max(anchor, focus));
  const view = doc.defaultView;
  if (!view) {
    settled(Math.min(anchor, focus), Math.max(anchor, focus));
    return true;
  }
  const move = (next: PointerEvent) => {
    if (next.pointerId !== event.pointerId) return;
    const offset = offsetAt(next.clientX, next.clientY);
    if (offset === null || offset === focus) return;
    focus = offset;
    select();
  };
  const release = (last: PointerEvent) => {
    if (last.pointerId !== event.pointerId) return;
    move(last);
    view.removeEventListener("pointermove", move, true);
    view.removeEventListener("pointerup", release, true);
    view.removeEventListener("pointercancel", release, true);
    anchor = Math.min(anchor, length);
    settled(Math.min(anchor, focus), Math.max(anchor, focus));
  };
  view.addEventListener("pointermove", move, true);
  view.addEventListener("pointerup", release, true);
  view.addEventListener("pointercancel", release, true);
  return true;
}

/** A triple-click's paragraph at `offset`: its words, without its break. */
export function presentationParagraphRange(text: string, offset: number): [number, number] {
  const at = Math.max(0, Math.min(text.length, Math.trunc(offset)));
  const breakAt = text.indexOf("\n", at);
  return [text.lastIndexOf("\n", at - 1) + 1, breakAt < 0 ? text.length : breakAt];
}

/**
 * A double- or triple-click on a drawn body being typed in (a text box's or a
 * shape's words): the word or the paragraph under the pointer, measured in
 * the body's own words (`presentationDoubleClickWord`), instead of the
 * browser's own word selection, which counts the drawn bullets and the rows'
 * breaks as text. Returns the model range it selected, or null when the press
 * is not one (a single click, another button) or misses the words; the caller
 * then lets the browser have it.
 */
export function selectPresentationPreviewWordAtPress(
  root: HTMLElement,
  event: MouseEvent,
  paragraphs?: readonly SlideParagraphWithRuns[],
): PresentationPreviewSelection | null {
  if (event.button !== 0 || event.detail < 2 || presentationPreviewComposing(root)) return null;
  const text = presentationPreviewPlainText(root);
  if (!text) return null;
  const hit = presentationPreviewOffsetAtPoint(root, event.clientX, event.clientY, text.length, paragraphs);
  if (hit === null) return null;
  const [start, end] = event.detail === 2 ? presentationDoubleClickWord(text, hit) : presentationParagraphRange(text, hit);
  event.preventDefault();
  if (root.ownerDocument.activeElement !== root) root.focus({ preventScroll: true });
  return setPresentationPreviewSelection(root, start, end) ? { start, end } : null;
}

/**
 * Where a caret at `offset` in `before` stands in `after`: unchanged before
 * the edited span, moved with the text after it, and at the end of the new
 * span when it stood inside the replaced one. A caret keeps to the person's
 * own typing while someone else edits the same text.
 */
export function mapOffsetThroughTextChange(before: string, after: string, offset: number): number {
  const limit = Math.min(before.length, after.length);
  let prefix = 0;
  while (prefix < limit && before.charCodeAt(prefix) === after.charCodeAt(prefix)) prefix += 1;
  let suffix = 0;
  while (
    suffix < limit - prefix &&
    before.charCodeAt(before.length - 1 - suffix) === after.charCodeAt(after.length - 1 - suffix)
  ) suffix += 1;
  const clamped = Math.max(0, Math.min(before.length, Math.trunc(offset)));
  if (clamped <= prefix) return clamped;
  if (clamped >= before.length - suffix) return clamped + (after.length - before.length);
  return after.length - suffix;
}

/** Map the visible rich-preview selection back to the block's flat text model. */
export function presentationPreviewSelection(
  root: HTMLElement,
  textLength: number,
  paragraphs?: readonly SlideParagraphWithRuns[],
): PresentationPreviewSelection | null {
  const selection = root.ownerDocument.defaultView?.getSelection();
  if (!selection || selection.rangeCount === 0) return null;
  const range = selection.getRangeAt(0);
  if (
    !root.contains(range.startContainer) ||
    !root.contains(range.endContainer)
  ) {
    return null;
  }
  const clamp = (value: number) =>
    Math.max(0, Math.min(textLength, Math.trunc(value)));
  const start = clamp(
    presentationPreviewPointOffset(
      root,
      range.startContainer,
      range.startOffset,
      paragraphs,
    ),
  );
  const end = clamp(
    presentationPreviewPointOffset(
      root,
      range.endContainer,
      range.endOffset,
      paragraphs,
    ),
  );
  return start <= end ? { start, end } : { start: end, end: start };
}

/**
 * Whether a delete the browser is about to make in a drawn text body takes
 * no words: the body is empty, or a caret stands at the words' start (a
 * backward delete) or their end (a forward one). The browser still acts
 * there -- it removes the empty row the body is drawn in. WebKit's Backspace
 * in an empty title took its row, so the caret fell to the top of the box,
 * out of the row its anchor places at the bottom, and the next key was typed
 * up there before the redraw put the words back down. Nothing in the words
 * changes, so the delete is refused before the browser makes it.
 */
export function presentationPreviewDeleteTakesNothing(
  root: HTMLElement,
  inputType: string | undefined,
  paragraphs?: readonly SlideParagraphWithRuns[],
): boolean {
  if (!inputType?.startsWith("delete")) return false;
  const textLength = presentationPreviewPlainText(root).length;
  if (textLength === 0) return true;
  const range = presentationPreviewSelection(root, textLength, paragraphs);
  if (!range || range.start !== range.end) return false;
  if (inputType.endsWith("Backward")) return range.start === 0;
  if (inputType.endsWith("Forward")) return range.end === textLength;
  return false;
}

const guardedPreviewBodies = new WeakSet<HTMLElement>();
/** Each guarded body's own writer of a replaced range (its latest). */
const previewSelectionWriters = new WeakMap<HTMLElement, (host: HTMLElement, inserted: string) => void>();

/** The paragraph's words a DOM point stands in, in a body drawn as rows;
 * null on a drawn bullet or number, or between the rows. */
function paragraphWordsAt(root: HTMLElement, node: Node): Element | null {
  const element = node.nodeType === Node.ELEMENT_NODE ? (node as Element) : node.parentElement;
  if (!element || element.closest("[data-presentation-selection-ignore]")) return null;
  const words = element.closest('[data-presentation-paragraph-text="true"]');
  return words && root.contains(words) ? words : null;
}



















export function presentationPreviewStructuralEdit(
  root: HTMLElement,
  inputType: string,
  data: string | null,
  ranges: readonly AbstractRange[] = [],
): { start: number; end: number; inserted: string } | null {
  const typed = inputType === "insertText" || inputType === "insertReplacementText";
  if (!typed && !inputType.startsWith("delete")) return null;
  if (!root.querySelector("[data-presentation-paragraph-index]")) return null;
  let range: AbstractRange | null = ranges[0] ?? null;
  if (!range) {
    const selection = root.ownerDocument.defaultView?.getSelection();
    range = selection && selection.rangeCount > 0 ? selection.getRangeAt(0) : null;
  }
  if (!range || !root.contains(range.startContainer) || !root.contains(range.endContainer)) return null;
  const startWords = paragraphWordsAt(root, range.startContainer);
  if (startWords && startWords === paragraphWordsAt(root, range.endContainer)) return null;
  const textLength = presentationPreviewPlainText(root).length;
  const clamp = (value: number) => Math.max(0, Math.min(textLength, Math.trunc(value)));
  const from = clamp(presentationPreviewPointOffset(root, range.startContainer, range.startOffset, undefined));
  const to = clamp(presentationPreviewPointOffset(root, range.endContainer, range.endOffset, undefined));
  const start = Math.min(from, to);
  const end = Math.max(from, to);
  // A delete with nothing in its range is the browser's to refuse or not.
  if (!typed && start === end) return null;
  return { start, end, inserted: typed ? (data ?? "") : "" };
}

/**
 * The native `beforeinput` guard of a drawn text body being typed in: a
 * delete with no words to take, and the browser's own formatting (a menu's
 * Bold, a `formatItalic`), which would change only the DOM, are refused.
 * React's `onBeforeInput` is not this event -- it is built from key presses
 * and text input, never fires for a delete and has no `inputType` -- so the
 * guards written there never ran. Once per host element: a host remounts
 * with its words (keyed by them), and the new one is guarded on mount.
 */
export function guardPresentationPreviewInput(
  host: HTMLElement | null,
  /**
   * Writes `inserted` over the body's selected words in the model (the host's
   * own replace). Given, a typed or deleted range the browser would edit
   * across the drawn rows is put through it instead
   * (`presentationPreviewStructuralEdit`).
   */
  replaceSelection?: (host: HTMLElement, inserted: string) => void,
): void {
  if (!host) return;
  if (replaceSelection) previewSelectionWriters.set(host, replaceSelection);
  if (guardedPreviewBodies.has(host)) return;
  guardedPreviewBodies.add(host);
  host.addEventListener("beforeinput", (event) => {
    if (event.isComposing || event.inputType === "deleteCompositionText") return;
    if (
      event.inputType.startsWith("format") ||
      presentationPreviewDeleteTakesNothing(host, event.inputType)
    ) {
      event.preventDefault();
      return;
    }
    const write = previewSelectionWriters.get(host);
    if (!write) return;
    const edit = presentationPreviewStructuralEdit(
      host,
      event.inputType,
      event.data ?? event.dataTransfer?.getData("text/plain") ?? null,
      typeof event.getTargetRanges === "function" ? event.getTargetRanges() : [],
    );
    if (!edit) return;
    event.preventDefault();
    setPresentationPreviewSelection(host, edit.start, edit.end);
    write(host, edit.inserted);
  });
  // The key itself as well: a WebKit build that raises no `beforeinput` for
  // a key's delete (only the `input` after it) still took the empty row.
  host.addEventListener("keydown", (event) => {
    if (event.isComposing || event.keyCode === 229) return;
    if (event.key !== "Backspace" && event.key !== "Delete") return;
    const inputType = event.key === "Backspace" ? "deleteContentBackward" : "deleteContentForward";
    if (presentationPreviewDeleteTakesNothing(host, inputType)) event.preventDefault();
  });
}

/** Read editable preview text without its visual bullet / number gutters. */
export function presentationPreviewPlainText(root: HTMLElement): string {
  const paragraphs = Array.from(
    root.querySelectorAll<HTMLElement>(
      "[data-presentation-paragraph-index]",
    ),
  );
  if (paragraphs.length === 0) return liveTextOf(root);
  return paragraphs
    .map((paragraph) =>
      paragraphRowText(
        paragraph.querySelector<HTMLElement>(
          '[data-presentation-paragraph-text="true"]',
        ),
      ),
    )
    .join("\n");
}

/** Resolve a flat model offset back into the DOM that paints the glyphs. */
export function presentationPreviewDomPoint(
  root: HTMLElement,
  rawOffset: number,
): PresentationDomPoint {
  // An equation atom counts as its characters and is never entered: an
  // offset inside it lands after it (`presentationAtomAwareDomPoint`).
  const textPoint = (
    textRoot: HTMLElement,
    rawLocal: number,
  ): PresentationDomPoint => presentationAtomAwareDomPoint(textRoot, rawLocal);

  let offset = Math.max(0, Math.trunc(rawOffset));
  const paragraphs = Array.from(
    root.querySelectorAll<HTMLElement>(
      "[data-presentation-paragraph-index]",
    ),
  );
  for (let index = 0; index < paragraphs.length; index += 1) {
    const textRoot = paragraphs[index]?.querySelector<HTMLElement>(
      '[data-presentation-paragraph-text="true"]',
    );
    if (!textRoot) continue;
    const text = paragraphRowText(textRoot);
    if (offset <= text.length || index === paragraphs.length - 1) {
      return textPoint(textRoot, Math.min(offset, text.length));
    }
    offset -= text.length + 1;
  }
  return textPoint(root, offset);
}

/** Restore the native caret/range inside the same DOM that paints the text. */
/**
 * The text body an input method is composing in, per document (set by the
 * composition events themselves, so every caller sees the same truth).
 */
const composingIn = new WeakMap<Document, Element | null>();
function watchComposition(doc: Document): void {
  if (composingIn.has(doc)) return;
  composingIn.set(doc, null);
  doc.addEventListener("compositionstart", (event) => {
    composingIn.set(doc, event.target instanceof Element ? event.target : null);
  }, true);
  doc.addEventListener("compositionend", () => composingIn.set(doc, null), true);
}

/** Whether an input method is composing in `root` (its marked text is the browser's). */
export function presentationPreviewComposing(root: HTMLElement): boolean {
  watchComposition(root.ownerDocument);
  const target = composingIn.get(root.ownerDocument);
  return Boolean(target && (root.contains(target) || target.contains(root)));
}

export function setPresentationPreviewSelection(
  root: HTMLElement,
  start: number,
  end: number,
): boolean {
  // Under an input method's marked text the caret is the browser's: a
  // selection written there commits the letters typed so far (pinyin became
  
  // where the composition already is.
  if (presentationPreviewComposing(root)) return true;
  const selection = root.ownerDocument.defaultView?.getSelection();
  if (!selection) return false;
  const textLength = presentationPreviewPlainText(root).length;
  const safeStart = Math.max(0, Math.min(textLength, Math.trunc(start)));
  const safeEnd = Math.max(0, Math.min(textLength, Math.trunc(end)));
  const from = presentationPreviewDomPoint(root, Math.min(safeStart, safeEnd));
  const to = presentationPreviewDomPoint(root, Math.max(safeStart, safeEnd));
  try {
    const range = root.ownerDocument.createRange();
    range.setStart(from.node, from.offset);
    range.setEnd(to.node, to.offset);
    selection.removeAllRanges();
    selection.addRange(range);
    return true;
  } catch {
    return false;
  }
}









export function presentationPreviewLineEdge(
  root: HTMLElement,
  offset: number,
  edge: "start" | "end",
): number | null {
  const text = presentationPreviewPlainText(root);
  const at = Math.max(0, Math.min(text.length, Math.trunc(offset)));
  const paragraphStart = text.lastIndexOf("\n", at - 1) + 1;
  const paragraphEnd = text.indexOf("\n", at) === -1 ? text.length : text.indexOf("\n", at);
  const boxAt = (place: number) => {
    const point = presentationPreviewDomPoint(root, place);
    try {
      const range = root.ownerDocument.createRange();
      range.setStart(point.node, point.offset);
      return officeCaretBox(range);
    } catch {
      return null;
    }
  };
  const here = boxAt(at);
  if (!here) return edge === "start" ? paragraphStart : paragraphEnd;
  const middle = here.top + here.height / 2;
  const sameLine = (place: number) => {
    const box = boxAt(place);
    return box !== null && box.top <= middle && box.top + box.height >= middle;
  };
  let place = at;
  if (edge === "start") {
    while (place > paragraphStart && sameLine(place - 1)) place -= 1;
  } else {
    while (place < paragraphEnd && sameLine(place + 1)) place += 1;
  }
  return place;
}

/**
 * Each equation's place in the flat text, `[start, end)`: an arrow key steps
 * over an equation whole. After an equation the caret stands in its
 * zero-width anchor, and the browser's Left took a press to cross that
 * character and only the next one to pass the equation.
 */
export function presentationPreviewAtomSpans(root: HTMLElement): Array<{ start: number; end: number }> {
  const spans: Array<{ start: number; end: number }> = [];
  for (const atom of Array.from(root.querySelectorAll<HTMLElement>("[data-presentation-math]"))) {
    const parent = atom.parentNode;
    if (!parent) continue;
    const index = Array.prototype.indexOf.call(parent.childNodes, atom);
    const start = presentationPreviewPointOffset(root, parent, index, undefined);
    const length = (atom.getAttribute("data-presentation-math-text") ?? "").length;
    spans.push({ start, end: start + length });
  }
  return spans;
}

