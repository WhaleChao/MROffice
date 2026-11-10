/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export const PDF_NATIVE_TEXT_LINE_ID_ATTR = "data-pdf-native-text-line-id";
const EDIT_BOX = "data-pdf-native-text-edit-box";
const STAGED_TEXT = "data-pdf-native-text-staged-text";
/** Words on their way to this box's line from an edit further back in the queue:
 * another line's revision came first, and the box was rebuilt as it stood. Only
 * for the line's identity here: the box does not read as staged (a press on it
 * asks the line model what the line reads). */
const EXPECTED_WORDS = "data-pdf-native-text-line-expects";
/** Where on its page a move being written puts this box's line: "x,y,width,
 * height" in page fractions, then the time (ms) the mark is good until. */
const GOES_TO = "data-pdf-native-text-line-goes-to";
/** A move is written, or refused, well inside this. */
const GOES_TO_FOR_MS = 120_000;

type Frame = Readonly<{ x: number; y: number; width: number; height: number }>;
/** `words`: what the line is expected to read after the rebuild (the words
 * staged into it, when an edit was); `shown`: what it read before;
 * `movedTo`: where a move being written puts it, `movedUntil` as marked. */
type LineRecord = Readonly<{ id: string; words: string; shown: string; staged: boolean; stagedWords: string; frame: Frame;
  movedTo: Frame | null; movedUntil: number }>;

/** What a page's layer held when it was last emptied. */
const pending = new WeakMap<HTMLElement, Readonly<{ lines: readonly LineRecord[]; at: number }>>();
/** Kept for the rebuild that follows (a slow machine starts it seconds later). */
const PENDING_FOR_MS = 60_000;

let sequence = 0;
const newLineId = () => `line-${Date.now().toString(36)}-${(sequence += 1).toString(36)}`;

const words = (text: string | null | undefined) => (text ?? "").replace(/\s+/g, " ").trim();

function frameOn(box: Element, page: Element): Frame | null {
  const at = box.getBoundingClientRect(), on = page.getBoundingClientRect();
  if (!(at.width > 0 && at.height > 0 && on.width > 0 && on.height > 0)) return null;
  return { x: (at.left - on.left) / on.width, y: (at.top - on.top) / on.height,
    width: at.width / on.width, height: at.height / on.height };
}

/** The same frame, within the rounding of a layout at another moment. */
const sameFrame = (left: Frame, right: Frame) =>
  Math.abs(left.x - right.x) <= 0.002 && Math.abs(left.y - right.y) <= 0.002 &&
  Math.abs(left.width - right.width) <= 0.004 && Math.abs(left.height - right.height) <= 0.004;

/** On one row and overlapping along it: the line an edit widened or
 * narrowed (its words changed, its start kept). */
const sameRow = (left: Frame, right: Frame) => {
  const leftMiddle = left.y + left.height / 2, rightMiddle = right.y + right.height / 2;
  return leftMiddle >= right.y && leftMiddle <= right.y + right.height &&
    rightMiddle >= left.y && rightMiddle <= left.y + left.height &&
    Math.min(left.x + left.width, right.x + right.width) > Math.max(left.x, right.x);
};

/** How far apart two frames stand (page fractions, place and size). */
const apart = (left: Frame, right: Frame) =>
  Math.abs(left.x - right.x) + Math.abs(left.y - right.y) +
  Math.abs(left.width - right.width) + Math.abs(left.height - right.height);

function readGoesTo(box: Element): Readonly<{ frame: Frame; until: number }> | null {
  const marked = box.getAttribute(GOES_TO);
  if (!marked) return null;
  const [x, y, width, height, until] = marked.split(",").map(Number);
  if (![x, y, width, height, until].every(Number.isFinite) || !(width > 0 && height > 0) || Date.now() >= until) return null;
  return { frame: { x, y, width, height }, until };
}

/**
 * Says where the move being written puts the line of `box` (its frame there,
 * in page fractions): the box that reads its words there on the revision
 * continues it. Null takes the mark back (the move was refused).
 */
export function markPdfNativeTextLineGoesTo(box: Element | null | undefined, frame: Frame | null): void {
  if (!box) return;
  if (!frame || ![frame.x, frame.y, frame.width, frame.height].every(Number.isFinite) ||
    !(frame.width > 0 && frame.height > 0)) { box.removeAttribute(GOES_TO); return; }
  box.setAttribute(GOES_TO, [frame.x, frame.y, frame.width, frame.height, Date.now() + GOES_TO_FOR_MS].join(","));
}

/** Whether a move being written puts the line of `box` somewhere else: the
 * box stands where the line started, its move's revision not on the page yet. */
export function pdfNativeTextLineIsGoing(box: Element | null | undefined): boolean {
  return Boolean(box && readGoesTo(box));
}

/** The page element a layer belongs to (the key the rebuild finds it by). */
const pageOf = (layer: Element) => layer.closest<HTMLElement>(".pdf-viewer__page");

/** Reads the boxes of a page's layer about to be emptied for a rebuild. */
export function recordPdfNativeTextLineLineage(layer: HTMLElement): void {
  const page = pageOf(layer);
  if (!page) return;
  const lines: LineRecord[] = [];
  for (const box of layer.querySelectorAll<HTMLElement>(`span[${EDIT_BOX}="true"][${PDF_NATIVE_TEXT_LINE_ID_ATTR}]`)) {
    const id = box.getAttribute(PDF_NATIVE_TEXT_LINE_ID_ATTR);
    const frame = frameOn(box, page);
    if (!id || !frame) continue;
    const staged = box.getAttribute(STAGED_TEXT) ?? box.getAttribute(EXPECTED_WORDS);
    const goesTo = readGoesTo(box);
    lines.push({ id, words: words(staged ?? box.textContent), shown: words(box.textContent), staged: staged !== null,
      stagedWords: staged ?? "", frame, movedTo: goesTo?.frame ?? null, movedUntil: goesTo?.until ?? 0 });
  }
  // A rebuild superseded before its boxes were written leaves the first
  // record: the layer it would read now is empty.
  if (lines.length > 0 || !pending.has(page)) pending.set(page, { lines, at: Date.now() });
}

/**
 * Gives each edit box of a page's layer its line id: the box it continues
 * (from the record taken when the layer was emptied), or a new one. A box
 * that already has one keeps it. Each id goes to one box.
 */
export function assignPdfNativeTextLineIds(page: HTMLElement): void {
  const layer = page.querySelector<HTMLElement>("[data-pdf-text-layer]") ?? page;
  const boxes = [...layer.querySelectorAll<HTMLElement>(`span[${EDIT_BOX}="true"]`)];
  if (boxes.length === 0) return;
  const record = pending.get(page);
  const lines = record && Date.now() - record.at < PENDING_FOR_MS ? record.lines : [];
  const taken = new Set<string>();
  for (const box of boxes) {
    const id = box.getAttribute(PDF_NATIVE_TEXT_LINE_ID_ATTR);
    if (id) taken.add(id);
  }
  const unnamed = boxes.filter(box => !box.hasAttribute(PDF_NATIVE_TEXT_LINE_ID_ATTR));
  if (unnamed.length === 0) return;
  // Nothing to continue (the page's first build): new ids, no layout read.
  if (lines.length === 0) {
    for (const box of unnamed) box.setAttribute(PDF_NATIVE_TEXT_LINE_ID_ATTR, newLineId());
    return;
  }
  const frames = new Map(unnamed.map(box => [box, frameOn(box, page)] as const));
  const continues = (line: LineRecord, box: HTMLElement) => {
    const frame = frames.get(box);
    if (!frame) return false;
    const read = words(box.textContent);
    // Its words where its move puts it; the staged words on its row; or (the
    // edit refused, undone, or another line's revision first) the words it
    // showed, where it stood.
    return (line.movedTo !== null && read === line.words && sameRow(line.movedTo, frame)) ||
      (line.staged && read === line.words && sameRow(line.frame, frame)) ||
      (read === line.shown && sameFrame(line.frame, frame));
  };
  // The lines being moved first, then the lines an edit was staged into:
  // their place or their words are new, and an unchanged line with the same
  // words elsewhere never takes their id.
  const ordered = [...lines.filter(line => line.movedTo !== null),
    ...lines.filter(line => line.movedTo === null && line.staged),
    ...lines.filter(line => line.movedTo === null && !line.staged)];
  const named = new Set<HTMLElement>();
  for (const line of ordered) {
    if (taken.has(line.id)) continue;
    let matches = unnamed.filter(box => !named.has(box) && continues(line, box));
    if (matches.length > 1 && line.movedTo) {
      // A twin of its words on the row it goes to: the one where it lands.
      const lands = line.movedTo;
      const there = matches.filter(box => apart(frames.get(box)!, lands) <= 0.01);
      if (there.length === 1) matches = there;
    }
    if (matches.length !== 1) continue;
    matches[0].setAttribute(PDF_NATIVE_TEXT_LINE_ID_ATTR, line.id);
    // Another line's revision came first, and this one's words are still on
    // their way. Forgotten here, the revision that brings them read as another
    // line, with a new id, and the edit queued behind it looked for the old one
    
    if (line.staged && words(matches[0].textContent) !== line.words) {
      matches[0].setAttribute(EXPECTED_WORDS, line.stagedWords);
    }
    // The same of a move: the box still stands where the line stood (nearer
    // to it than to where the move puts it), its move's revision yet to come.
    const stands = frames.get(matches[0]);
    if (line.movedTo && stands && apart(stands, line.frame) < apart(stands, line.movedTo) && Date.now() < line.movedUntil) {
      matches[0].setAttribute(GOES_TO,
        [line.movedTo.x, line.movedTo.y, line.movedTo.width, line.movedTo.height, line.movedUntil].join(","));
    }
    named.add(matches[0]);
    taken.add(line.id);
  }
  // A line with another edit left on it before this revision came (queued
  // behind it): the revision reads the edit before, not the words its box was
  // last marked with. Still that line: the one box left on its row. Named
  // otherwise, what was following the line lost it, and the edit queued on it
  
  for (const line of lines) {
    if (!line.staged || taken.has(line.id)) continue;
    const onRow = unnamed.filter(box => {
      const frame = frames.get(box);
      return !named.has(box) && frame && sameRow(line.frame, frame);
    });
    if (onRow.length !== 1) continue;
    onRow[0].setAttribute(PDF_NATIVE_TEXT_LINE_ID_ATTR, line.id);
    // Its words are still on their way to it.
    onRow[0].setAttribute(STAGED_TEXT, line.stagedWords);
    named.add(onRow[0]);
    taken.add(line.id);
  }
  for (const box of unnamed) {
    // (A box with no frame continues no line, but is named all the same:
    // what follows a line waits for every box to be named.)
    if (!named.has(box)) box.setAttribute(PDF_NATIVE_TEXT_LINE_ID_ATTR, newLineId());
  }
  // The record stays until the layer is emptied again (or it is old): a
  // layer whose boxes come in more than one pass names the later ones from
  // it too; an id already given is never given twice.
}

/** The line id of a box (or of a held picture's box). */
export function pdfNativeTextLineIdOf(box: Element | null | undefined): string | null {
  return box?.getAttribute(PDF_NATIVE_TEXT_LINE_ID_ATTR) ?? null;
}

/** The page's edit box of a line, when it is there now. */
export function findPdfNativeTextBoxByLineId(page: ParentNode | null | undefined, lineId: string | null): HTMLElement | null {
  if (!page || !lineId) return null;
  const escaped = typeof CSS !== "undefined" && CSS.escape ? CSS.escape(lineId) : lineId.replace(/"/g, '\\"');
  const found = [...page.querySelectorAll<HTMLElement>(`span[${EDIT_BOX}="true"][${PDF_NATIVE_TEXT_LINE_ID_ATTR}="${escaped}"]`)];
  return found.length === 1 ? found[0] : null;
}

/** What a held picture of the layer keeps of a box (it is paint only, but a
 * press on it names its line). */
export const PDF_NATIVE_TEXT_GHOST_KEPT_ATTRS: readonly string[] = [PDF_NATIVE_TEXT_LINE_ID_ATTR, STAGED_TEXT];
