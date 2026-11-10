/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { pdfNativeTextEditFrameClientRect } from "./native-text-edit-frame-rect";

/**
 * A click on a page whose text layer is being rebuilt (a saved revision just
 * replaced the page: no editable boxes yet, or boxes not yet interactive) used
 * to fall through: the reader saw nothing happen and had to click again. The
 * click is kept as an intent -- its page and its point as a fraction of that
 * page -- and the box under that point is activated as soon as the layer is
 * back, through the ordinary activation (which rebinds the current bytes,
 * proof and authority; the intent grants nothing).
 *
 * The intent ends at the first activation that is taken, at any newer pointer
 * gesture, or at its deadline.
 */
export type PdfNativeTextRebuildClickActivate = (
  target: HTMLElement,
  point: Readonly<{ clientX: number; clientY: number }>,
  /** The kept click was the first of a double click: open with the caret,
   * as a double click on a ready box does, not as a block. */
  repeated: boolean,
) => boolean;

/** Cancels the kept click; `repeat` records the second press of the same
 * double click (the caller's own `detail > 1` click). */
export type PdfNativeTextRebuildClickHandle = (() => void) & { repeat: () => void };

/** A second press this close to the kept click, this soon, is the rest of
 * the same double click -- not a newer gesture. */
const REPEAT_WINDOW_MS = 500;
const REPEAT_DISTANCE_PX = 8;

const EDIT_BOX = 'span[data-pdf-native-text-edit-box="true"]';
const SELECTED = "data-pdf-native-text-edit-selected";
/** A box held as a picture while the layer is rebuilt (text-layer-ghost). */
const HELD_BOX = 'span[data-pdf-native-text-edit-box-ghost="true"]';

/** The editable box under a page-relative point: the smallest box whose
 * drawn frame holds it (overlays above the text layer do not hide it). The
 * frame, not the span: a span can sit away from the frame the reader sees
 * (pdfNativeTextEditFrameClientRect). */
export function findPdfNativeTextBoxAtPagePoint(
  page: HTMLElement,
  fraction: Readonly<{ x: number; y: number }>,
): { span: HTMLElement; point: { clientX: number; clientY: number } } | null {
  const rect = page.getBoundingClientRect();
  if (!(rect.width > 0 && rect.height > 0)) return null;
  const point = { clientX: rect.left + rect.width * fraction.x, clientY: rect.top + rect.height * fraction.y };
  let best: HTMLElement | null = null;
  let bestArea = Number.POSITIVE_INFINITY;
  for (const span of page.querySelectorAll<HTMLElement>(EDIT_BOX)) {
    if (span.getAttribute("data-pdf-native-text-editable") !== "true") continue;
    const box = pdfNativeTextEditFrameClientRect(span);
    if (point.clientX < box.left || point.clientX > box.right || point.clientY < box.top || point.clientY > box.bottom) continue;
    const area = (box.right - box.left) * (box.bottom - box.top);
    if (area > 0 && area < bestArea) { best = span; bestArea = area; }
  }
  return best ? { span: best, point } : null;
}

export function queuePdfNativeTextRebuildClick(input: Readonly<{
  root: HTMLElement;
  pageNumber: number;
  clientX: number;
  clientY: number;
  activate: PdfNativeTextRebuildClickActivate;
  deadlineMs?: number;
  onSettled?: (outcome: "activated" | "expired" | "superseded") => void;
  /** Whether what the reader did since makes this click out of date. */
  isSuperseded?: () => boolean;
}>): PdfNativeTextRebuildClickHandle {
  const findPage = () => input.root.querySelector<HTMLElement>(`.pdf-viewer__page[data-page-number="${input.pageNumber}"]`);
  const page = findPage();
  const rect = page?.getBoundingClientRect();
  if (!page || !rect || !(rect.width > 0 && rect.height > 0)) {
    return Object.assign(() => undefined, { repeat: () => undefined });
  }
  const keptAt = Date.now();
  let repeated = false;
  const fraction = { x: (input.clientX - rect.left) / rect.width, y: (input.clientY - rect.top) / rect.height };
  let frame = 0;
  let done = false;
  // The click is kept, and says so: the box under it is shown selected at
  // once, in the layer that is still on the page. Until the rebuilt layer
  // answered, nothing did for most of a second after a save landed.
  let marked: HTMLElement | null = null;
  let markedArea = Number.POSITIVE_INFINITY;
  for (const span of page.querySelectorAll<HTMLElement>(`${EDIT_BOX}, ${HELD_BOX}`)) {
    const box = pdfNativeTextEditFrameClientRect(span);
    if (input.clientX < box.left || input.clientX > box.right || input.clientY < box.top || input.clientY > box.bottom) continue;
    const area = (box.right - box.left) * (box.bottom - box.top);
    if (area > 0 && area < markedArea) { marked = span; markedArea = area; }
  }
  if (marked?.hasAttribute(SELECTED)) marked = null;
  marked?.setAttribute(SELECTED, "true");
  const observer = typeof MutationObserver === "function" ? new MutationObserver(() => schedule()) : null;
  const finish = (outcome: "activated" | "expired" | "superseded") => {
    if (done) return;
    done = true;
    observer?.disconnect();
    if (frame) cancelAnimationFrame(frame);
    clearTimeout(timer);
    clearTimeout(poll);
    document.removeEventListener("pointerdown", supersede, true);
    // The activation marks the box it opens; a box left in the old layer,
    // or one whose click came to nothing, is not left looking selected.
    if (marked && marked !== activated) marked.removeAttribute(SELECTED);
    input.onSettled?.(outcome);
  };
  let activated: HTMLElement | null = null;
  const attempt = () => {
    frame = 0;
    if (done) return;
    if (input.isSuperseded?.()) { finish("superseded"); return; }
    const current = findPage();
    const hit = current ? findPdfNativeTextBoxAtPagePoint(current, fraction) : null;
    if (hit && input.activate(hit.span, hit.point, repeated)) { activated = hit.span; finish("activated"); return; }
    // Interactivity can return without a DOM change this observer sees.
    clearTimeout(poll);
    poll = setTimeout(schedule, 100);
  };
  let poll: ReturnType<typeof setTimeout> | undefined;
  function schedule() { if (!done && !frame) frame = requestAnimationFrame(attempt); }
  const supersede = (event: Event) => {
    // A human double click presses again a few hundred milliseconds later:
    // that press used to drop the kept click, and the double click on a
    // box still being prepared did nothing at all.
    if (
      Date.now() - keptAt <= REPEAT_WINDOW_MS &&
      "clientX" in event && "clientY" in event &&
      Math.hypot(Number(event.clientX) - input.clientX, Number(event.clientY) - input.clientY) <= REPEAT_DISTANCE_PX
    ) {
      repeated = true;
      return;
    }
    finish("superseded");
  };
  const timer = setTimeout(() => finish("expired"), input.deadlineMs ?? 4000);
  observer?.observe(input.root, {
    childList: true, subtree: true, attributes: true,
    attributeFilter: ["data-pdf-native-text-editable", "data-pdf-native-text-edit-box"],
  });
  // The click that queued this intent is still being dispatched: a newer
  // gesture is only one that starts after it.
  setTimeout(() => document.addEventListener("pointerdown", supersede, true), 0);
  schedule();
  return Object.assign(() => finish("superseded"), { repeat: () => { repeated = true; } });
}
