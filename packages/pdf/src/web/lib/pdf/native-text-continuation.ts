"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useEffect, useLayoutEffect, useRef, useState, type RefObject } from "react";
import type { PdfWorkingDisplaySource } from "./working-display-source";
import type { PdfNativeListTextSelectionEvidence } from "./engine/native/list-text-fidelity";
import { readPdfTextKeyboardInput } from "./text-keyboard-input";
import { markPdfTextLatency } from "./text-edit-latency";
import { reportOfficeExperience } from "@mrnote/core/web/lib/office-experience-report.ts";
import { findPdfNativeTextBoxByLineId, pdfNativeTextLineIdOf, pdfNativeTextLineIsGoing } from "./native-text-line-lineage";

/** User intent only. Completion grants no editing authority: the reader must
 * resolve the target in its current text layer and run normal activation.
 * The box opens selected as a block (WPS's single click) unless `caret`
 * says the press was the second of a double click: then with its caret. */
export type PdfNativeTextContinuationCapture = ((
  target: HTMLElement,
  point: Readonly<{ clientX: number; clientY: number }>,
  options?: Readonly<{ caret?: boolean }>,
) => ((accepted: boolean) => void) | null) & {
  /** The line left has let go of the page (its editor closed on staging an
   * edit): the line captured last opens queued now, in that same update,
   * when the page's boxes are held for the revision staged. */
  openQueuedNow?: () => void;
  /** The line opened queued was let go (its words handed to the window):
   * its taking-over is no longer waited for, and the line pressed next is
   * not held behind it. */
  dropAdoption?: () => void;
  /** The line open on this box lost the bytes it was being proved on (another
   * edit reached the working copy): it stays open, and is proved on the
   * revision that edit brings, found there by its id. */
  followLine?: (target: HTMLElement) => boolean;
  /** A line by its id on the revision the page shows (null: not there, or
   * the page between revisions). */
  readLine?: (lineId: string, pageNumber: number) => PdfNativeListTextSelectionEvidence | null;
};

const EDIT_BOX = "data-pdf-native-text-edit-box";
/** Set by the window on the span of a line whose words it just staged. */
export const PDF_NATIVE_TEXT_STAGED_TEXT_ATTR = "data-pdf-native-text-staged-text";
const STAGED_TEXT = PDF_NATIVE_TEXT_STAGED_TEXT_ATTR;
const PARAGRAPH_ID = "data-pdf-native-text-paragraph-id";
const ITEM_INDEX = "data-pdf-native-text-item-index";
const EDITABLE = "data-pdf-native-text-editable";
/** "queued": opened at once on the page as shown while its boxes are read
 * in again; "released": nothing to adopt (the queued box was closed). */
/** Dispatched on the reader's root with `{ target, text, typed, x, pageNumber }`
 * when keys typed for a line that never opened are given up. */
export const PDF_NATIVE_TEXT_ABANDONED_EVENT = "pdf-native-text-typeahead-abandoned";
/** Dispatched on a box's input when keys kept for it moved its caret (Home,
 * End) or typed into it: where its caret is is the reader's from then on --
 * the press that opened it no longer places it (PdfSourceTextCaret). */
export const PDF_NATIVE_TEXT_CARET_SETTLED_EVENT = "pdf-native-text-caret-settled";

export type PdfNativeTextActivationStatus = "accepted" | "transient-refusal" | "unsupported" | "queued" | "released";
type Outcome = PdfNativeTextActivationStatus;

/** A duplicate gesture joins the existing activation; it must not leave a
 * continuation waiting for a callback that only the first caller receives. */
export function joinPdfNativeTextActivation(
  pending: Readonly<{trigger:HTMLElement|null;selection?:Readonly<{pageNumber:number;sourceSha256:string}>;
    completion:Promise<Outcome>}> | null,
  target: HTMLElement, pageNumber: number, readCurrentSource: () => string | null,
  onSettled?: (status: Outcome) => void,
): boolean {
  if (!pending || pending.trigger!==target || pending.selection?.pageNumber!==pageNumber ||
      pending.selection.sourceSha256!==readCurrentSource()) return false;
  if (onSettled) void pending.completion.then(status=>{
    onSettled(target.isConnected && pending.selection?.sourceSha256===readCurrentSource() ? status : "transient-refusal");
  },()=>onSettled("transient-refusal"));
  return true;
}

type PageRelativeRect = Readonly<{x:number;y:number;width:number;height:number}>;
function pageRelativeRect(target: HTMLElement, page: HTMLElement | null): PageRelativeRect | null {
  if (!page) return null;
  const owner=page.getBoundingClientRect(),box=target.getBoundingClientRect();
  if (![owner.left,owner.top,owner.width,owner.height,box.left,box.top,box.width,box.height].every(Number.isFinite) ||
      owner.width<=0 || owner.height<=0 || box.width<=0 || box.height<=0) return null;
  return {x:(box.left-owner.left)/owner.width,y:(box.top-owner.top)/owner.height,
    width:box.width/owner.width,height:box.height/owner.height};
}
const INLINE_INPUT = "data-pdf-native-text-inline-input";
const TYPE_AHEAD_LIMIT = 200;
function isEditable(node: EventTarget | null): boolean {
  return node instanceof HTMLElement &&
    (node.isContentEditable || node.tagName === "INPUT" || node.tagName === "TEXTAREA" || node.tagName === "SELECT");
}
/** Types text at the box's caret the way the keyboard would: through the
 * element's own value setter and an input event, so the editor's state
 * follows. */
export function insertPdfNativeTextTypeAhead(input: HTMLTextAreaElement, text: string,
  /** Home or End pressed before the keys (or alone): the caret goes to the
   * line's start or end first, as the key would have put it. */
  move: "start" | "end" | null = null): void {
  if (move) {
    const at = move === "end" ? input.value.length : 0;
    input.setSelectionRange(at, at);
    input.dispatchEvent(new Event(PDF_NATIVE_TEXT_CARET_SETTLED_EVENT, { bubbles: true }));
    if (!text) return;
  }
  const start = input.selectionStart ?? input.value.length, end = input.selectionEnd ?? start;
  const value = input.value.slice(0, start) + text + input.value.slice(end);
  const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, "value")?.set;
  if (setter) setter.call(input, value); else input.value = value;
  input.dispatchEvent(new Event("input", { bubbles: true }));
  const caret = start + text.length;
  input.setSelectionRange(caret, caret);
}
function samePageRelativeRect(left: PageRelativeRect, right: PageRelativeRect | null): boolean {
  // Allow small CSS rounding differences across zoom and Ribbon layout. This
  // is target intent only; current activation still independently proves it.
  return right!==null && (["x","y","width","height"] as const).every(key=>Math.abs(left[key]-right[key])<=0.0005);
}

function relatedRegroupedTarget(
  original: PageRelativeRect, current: PageRelativeRect | null,
  originalText: string, currentText: string, point: Readonly<{x:number;y:number}>,
): boolean {
  if (!current || !originalText.trim() || !currentText.trim()) return false;
  const tolerance=0.0005;
  const contains=(a:PageRelativeRect,b:PageRelativeRect)=>b.x>=a.x-tolerance && b.y>=a.y-tolerance &&
    b.x+b.width<=a.x+a.width+tolerance && b.y+b.height<=a.y+a.height+tolerance;
  if (!((originalText.includes(currentText)&&contains(original,current)) ||
        (currentText.includes(originalText)&&contains(current,original)))) return false;
  const x=original.x+original.width*point.x,y=original.y+original.height*point.y;
  return x>=current.x-tolerance&&x<=current.x+current.width+tolerance&&
    y>=current.y-tolerance&&y<=current.y+current.height+tolerance;
}

/** The longest a press waits for the page's boxes held for the reader's own
 * edit going up (the line left let go; usePdfNativeTextContinuation). */
export const PRESS_WAIT_FOR_HELD_PAGE_MS = 60_000;

type ContinuationIntent = {
  dataItemId: string | undefined; pageNumber: number; sourceRevision: string | number;
  paragraphId: string | null; itemIndex: string | null; text: string;
  /** The line's id across the layer's rebuilds (native-text-line-lineage). */
  lineId: string | null;
  layoutId: string | null; x: number; y: number; ready: boolean; deadline: number; inFlight: boolean;
  /** When the line left let go (ready): the press waits for the page's boxes
   * from then, at most PRESS_WAIT_FOR_HELD_PAGE_MS while they are held. */
  readyAt?: number;
  /** The words an edit just staged into this line (the page still shows
   * the old ones): its box on the revision read in carries these. */
  stagedText: string | null;
  workingSessionId: string | null; workingRevision: number | null;
  pageRect: PageRelativeRect | null;
  typed: string;
  /** Home / End pressed before the box took the keyboard, ahead of `typed`. */
  caretMove: "start" | "end" | null;
  /** The press was the second of a double click: the caret goes in. */
  caret: boolean;
  /** The box opened: the intent is spent, but its keys still wait for
   * the input to take focus (until deliverUntil). */
  completed: boolean; deliverUntil: number;
  /** Opened at once on the page as shown while its boxes were read in
   * again; opened on them (adopting it) once they are back. */
  queued: boolean;
  queuedRefused: boolean;
  /** Its opening on the revision has looked for its line (said once). */
  lookedFor?: boolean;
  /** A press came after this one, before it was settled. */
  pressedOver?: boolean;
};

/** The intent's line in the page's current text layer, by its unchanged
 * page position (or, without one, its exact old identity), and where in it
 * the press was. Null while it is not there, or not editable now. */
function findPdfNativeTextIntentTarget(root: HTMLElement, intent: ContinuationIntent, allowPending: boolean,
  /** Taking over the box opened queued: the edit just staged may have moved
   * its line (a paragraph above set again); the one span with its text is
   * that line. */
  byTextAlone = false,
  /** Taking over a box, its revision up: the line is looked for once, here. */
  adopting = false) {
  const page = root.querySelector<HTMLElement>(`.pdf-viewer__page[data-page-number="${intent.pageNumber}"]`);
  const spans = [...(page?.querySelectorAll<HTMLElement>(`[data-pdf-text-layer] span[${EDIT_BOX}="true"]`) ?? [])];
  // The line by its id, carried over each rebuild of the layer
  // (native-text-line-lineage.ts): found whatever the edit just staged did
  // to its words, items or frame. The place-and-words search below is for
  // a box no id names (a layer built before it had them).
  const byLine = intent.lineId
    ? findPdfNativeTextBoxByLineId(page?.querySelector("[data-pdf-text-layer]"), intent.lineId) : null;
  // Found any other way, a line with an id says so: these searches go once the
  // live site never comes here (lib/office-experience-report, pdf_long_way).
  if (intent.lineId && !byLine && adopting) {
    reportOfficeExperience("pdf_long_way", { pageId: intent.dataItemId ?? "", pageType: "pdf", code: "line-not-by-id" });
    if (process.env.NODE_ENV !== "production") markPdfTextLatency("open.line-not-by-id", { lineId: intent.lineId, text: intent.text.slice(0, 24),
      staged: intent.stagedText?.slice(0, 24) ?? null, layers: page?.querySelectorAll("[data-pdf-text-layer]").length ?? 0,
      boxes: spans.map(span => `${pdfNativeTextLineIdOf(span)}:${(span.textContent ?? "").slice(0, 10)}`).slice(0, 14) });
  }
  if (!intent.lineId && adopting) {
    reportOfficeExperience("pdf_long_way", { pageId: intent.dataItemId ?? "", pageType: "pdf", code: "line-no-id" });
    markPdfTextLatency("open.line-no-id");
  }
  let matches = byLine ? [byLine] : spans.filter(span => {
    const text=span.textContent??"";
    const layoutId=span.getAttribute("data-pdf-native-text-layout-id");
    const stableLayoutId=intent.layoutId && !intent.layoutId.startsWith("native-text:");
    if (stableLayoutId && layoutId!==intent.layoutId) return false;
    if (intent.pageRect) {
      // Deleting a preceding font slot renumbers PDF.js items/paragraphs.
      // Match the unchanged page position, then activate its CURRENT item.
      const rect=pageRelativeRect(span,page);
      return ((text===intent.text || (intent.stagedText !== null && text===intent.stagedText)) &&
        samePageRelativeRect(intent.pageRect,rect)) ||
        (!stableLayoutId&&relatedRegroupedTarget(intent.pageRect,rect,intent.text,text,intent));
    }
    // Without page geometry preserve only the old exact identity. Never
    // select the first text match or guess a new item ordinal.
    return text===intent.text && (intent.paragraphId ? span.getAttribute(PARAGRAPH_ID) === intent.paragraphId :
      span.getAttribute(ITEM_INDEX) === intent.itemIndex) && layoutId===intent.layoutId;
  });
  let moved = false;
  if (byLine && intent.pageRect) {
    // Its frame changed (new words): the press is placed in it as it was
    // in the old one, unless the old place is still inside it.
    const rect = pageRelativeRect(byLine, page);
    const x = intent.pageRect.x + intent.pageRect.width * intent.x, y = intent.pageRect.y + intent.pageRect.height * intent.y;
    moved = !rect || x < rect.x || x > rect.x + rect.width || y < rect.y || y > rect.y + rect.height;
  }
  if (!matches.length && byTextAlone && (intent.stagedText ?? intent.text).trim()) {
    // Spaces as the text layer joins the line's items, not as typed.
    const words = (text: string) => text.replace(/\s+/g, " ").trim();
    const wanted = words(intent.stagedText ?? intent.text);
    const same = spans.filter(span => words(span.textContent ?? "") === wanted);
    if (same.length === 1) { matches = same; moved = true; }
  }
  // Opened queued, what was read of the line when it was pressed decides
  // (the viewer's snapshot): its box may be held meanwhile, "pending" while
  // the next revision is read in or "false" while the line left saves.
  if (matches.length !== 1 ||
    !(matches[0].getAttribute(EDITABLE) === "true" || (allowPending && matches[0].hasAttribute(EDITABLE)))) return null;
  const target = matches[0], rect = target.getBoundingClientRect();
  if (rect.width <= 0 || rect.height <= 0) return null;
  const pageBounds=page?.getBoundingClientRect();
  const point=!moved&&intent.pageRect&&pageBounds&&pageBounds.width>0&&pageBounds.height>0
    ? {clientX:pageBounds.left+pageBounds.width*(intent.pageRect.x+intent.pageRect.width*intent.x),
       clientY:pageBounds.top+pageBounds.height*(intent.pageRect.y+intent.pageRect.height*intent.y)}
    : {clientX:rect.left+rect.width*intent.x,clientY:rect.top+rect.height*intent.y};
  return { target, point };
}

export function usePdfNativeTextContinuation({
  rootRef, currentPage, dataItemId, sourceRevision, modeActive, writable,
  editingActive, interactive, activate, onCaptureChange,
  workingSource, queued, readLine,
}: {
  rootRef: RefObject<HTMLElement | null>;
  currentPage: number;
  dataItemId?: string;
  sourceRevision: string | number;
  workingSource?: PdfWorkingDisplaySource | null;
  modeActive: boolean;
  writable: boolean;
  editingActive: boolean;
  interactive: boolean;
  activate: (target: EventTarget | null, point: Readonly<{clientX: number; clientY: number}> | null,
    prepareOnly: boolean, onSettled: (status: Outcome) => void, selectObject?: boolean,
    /** "adopt": open the line on the revision read in, taking over the box
     * opened `queued` (its draft and caret). */
    mode?: "adopt") => boolean;
  onCaptureChange?: (capture: PdfNativeTextContinuationCapture | null) => void;
  readLine?: PdfNativeTextContinuationCapture["readLine"];
  /** Opening a line at once on the page as shown while its boxes wait for a
   * revision being read in (WPS switches lines instantly): `snapshot` reads,
   * when the line is pressed, what that needs (the document under it is let
   * go once the next revision starts loading); `open` opens it, "queued"
   * when it did. Without it, a line waits for its boxes. */
  queued?: Readonly<{
    snapshot: (target: HTMLElement, point: Readonly<{clientX: number; clientY: number}>, selectObject: boolean) => void;
    open: (target: HTMLElement, point: Readonly<{clientX: number; clientY: number}>, selectObject: boolean,
      onSettled: (status: Outcome) => void) => boolean;
  }>;
}) {
  const sourceDocumentRevisionRef = useRef(sourceRevision);
  const workingSourceRef = useRef(workingSource);
  const queuedRef = useRef(queued);
  const readLineRef = useRef(readLine);
  readLineRef.current = readLine;
  useLayoutEffect(() => { queuedRef.current = queued; }, [queued]);
  useLayoutEffect(() => {
    sourceDocumentRevisionRef.current = sourceRevision;
    workingSourceRef.current = workingSource;
  }, [sourceRevision, workingSource]);
  // A dirty editor consumes the pointer gesture while its exact draft saves.
  // Keep the target intent, not the old node/selection or a synthetic click.
  const [nativeTextContinuationEpoch, setNativeTextContinuationEpoch] = useState(0);
  const gestureCleanupRef = useRef<(() => void) | null>(null);
  const nativeTextContinuationRef = useRef<ContinuationIntent | null>(null);
  /** The line opened queued, until it is opened on the revision being read
   * in. Kept apart from the intent: a press meanwhile -- the caret moved in
   * that box, another line chosen -- leaves that box the reader's, and only
   * its opening on the revision can write it (a press used to drop it with
   * the intent, and the box waited for a proof for good). */
  const adoptionRef = useRef<ContinuationIntent | null>(null);
  const openQueuedNowRef = useRef<(intent: ContinuationIntent) => void>(() => undefined);
  openQueuedNowRef.current = (intent) => {
    const root = rootRef.current;
    const opening = queuedRef.current;
    if (!opening) return;
    if (!root || intent.inFlight || intent.queued || intent.queuedRefused || adoptionRef.current !== null) {
      markPdfTextLatency("open.queued-now-skipped", { inFlight: intent.inFlight, queued: intent.queued,
        refused: intent.queuedRefused, adopting: adoptionRef.current !== null });
      return;
    }
    const found = findPdfNativeTextIntentTarget(root, intent, true);
    if (!found) { markPdfTextLatency("open.queued-now-no-target"); return; }
    intent.inFlight = true;
    // Declined (no revision being read in, nothing read of the line), the
    // line opens as ever: nothing is marked refused here.
    if (!opening.open(found.target, found.point, !intent.caret, status => {
      intent.inFlight = false;
      if (status === "queued") queuedOpened(intent, found.target);
    })) intent.inFlight = false;
  };
  /** Open, and the reader's: the intent is spent (its keys go to that box);
   * its opening on the revision is followed apart (adoptionRef). */
  const queuedOpened = (intent: ContinuationIntent, target: HTMLElement) => {
    // The words the box opened with: the line just left may have staged an
    // edit into it after it was pressed (the press came while that edit was
    // still being compiled, its box not marked yet), and its box on the
    // revision carries those words, not the ones pressed. Unread, the box
    // was never found again, never taken over, and every key typed in it
    
    intent.stagedText = target.getAttribute(STAGED_TEXT) ?? intent.stagedText;
    intent.queued = true;
    intent.completed = true;
    intent.deliverUntil = Date.now() + 3_000;
    intent.deadline = Date.now() + 15_000;
    adoptionRef.current = intent;
    if (nativeTextContinuationRef.current === intent) nativeTextContinuationRef.current = null;
    setNativeTextContinuationEpoch(epoch => epoch + 1);
  };
  const stopTypeAheadRef = useRef<((reason?: "abandon") => void) | null>(null);
  useEffect(() => {
    const capture: PdfNativeTextContinuationCapture = (target, point, options) => {
      const root = rootRef.current;
      const page = target.closest<HTMLElement>(".pdf-viewer__page");
      const rect = target.getBoundingClientRect();
      // A box held as a picture while its layer is built again is the line
      // too: found again by its place and words once the layer is back.
      const held = target.getAttribute("data-pdf-native-text-edit-box-ghost") === "true";
      if (!root?.contains(target) || (target.getAttribute(EDIT_BOX) !== "true" && !held) ||
        Number(page?.dataset.pageNumber) !== currentPage || !dataItemId || rect.width <= 0 || rect.height <= 0) {
        if (process.env.NODE_ENV !== "production") markPdfTextLatency("open.capture-refused", {
          inRoot: Boolean(root?.contains(target)), box: target.getAttribute(EDIT_BOX), page: Number(page?.dataset.pageNumber),
          currentPage, item: Boolean(dataItemId), width: Math.round(rect.width), height: Math.round(rect.height) });
        return null;
      }
      const intent: ContinuationIntent = {
        dataItemId, pageNumber: currentPage, sourceRevision: sourceDocumentRevisionRef.current,
        paragraphId: target.getAttribute(PARAGRAPH_ID),
        itemIndex: target.getAttribute(ITEM_INDEX),
        layoutId: target.getAttribute("data-pdf-native-text-layout-id"), text: target.textContent ?? "",
        lineId: pdfNativeTextLineIdOf(target),
        stagedText: target.getAttribute(STAGED_TEXT),
        x: Math.max(0, Math.min(1, (point.clientX - rect.left) / rect.width)),
        y: Math.max(0, Math.min(1, (point.clientY - rect.top) / rect.height)),
        ready: false, deadline: 0, inFlight: false,
        pageRect: pageRelativeRect(target,page),
        workingSessionId: workingSourceRef.current?.sessionId ?? null,
        workingRevision: workingSourceRef.current?.workingRevision ?? null,
        typed: "", caretMove: null, caret: options?.caret === true, completed: false, deliverUntil: 0, queued: false, queuedRefused: false,
      };
      nativeTextContinuationRef.current = intent;
      // Read now, while the page shown is the one it was pressed on: if the
      // boxes are held for the next revision when this line is to open, it
      // opens queued from this, at once.
      try { queuedRef.current?.snapshot(target, point, !intent.caret); }
      catch { /* a queued opening is a speed-up only: the line opens anyway */ }
      // The line being left lets go of the keyboard at the press: a key
      // typed right after it is this box's (typed ahead below, or into this
      // box when it opens), never the old line's. Kept, it went into the old
      // line, changed the draft its finish was asked for, and the switch was
      
      const focused = document.activeElement;
      if (focused instanceof HTMLTextAreaElement && focused.getAttribute(INLINE_INPUT) === "true" &&
        root.contains(focused)) focused.blur();
      // Keys pressed before this box takes focus -- while the previous box
      // is still saving (its input is read-only meanwhile) and this one has
      // not opened -- went nowhere. They are kept, in order, and typed at
      // the box's caret the moment its input is focused; a key that reaches
      // a focused input is the input's own. Escape drops the intent.
      stopTypeAheadRef.current?.("abandon");
      let typeAheadFrame = 0;
      let laterPress: ((event: PointerEvent) => void) | null = null;
      const stopTypeAhead = (reason?: "abandon") => {
        // Keys typed for a line that never opened, given up for a newer press
        // (or its wait running out): they are the reader's words for that
        // line, handed to the window to write there (PDF_NATIVE_TEXT_ABANDONED_EVENT),
        // never dropped -- 26 keys were lost when the reader moved on while
        
        if (process.env.NODE_ENV !== "production") markPdfTextLatency("open.typeahead-stop", { reason: reason ?? null,
          keys: intent.typed.length, completed: intent.completed, current: nativeTextContinuationRef.current === intent });
        if (reason === "abandon" && intent.typed && !intent.completed) {
          const root = rootRef.current;
          // Found where it is now, or by where it was pressed on its page
          // (its boxes may be being built again right now).
          const found = root ? findPdfNativeTextIntentTarget(root, intent, true) : null;
          if (root && (found || intent.pageRect)) {
            if (process.env.NODE_ENV !== "production") markPdfTextLatency("open.typeahead-abandoned", { found: Boolean(found) });
            root.dispatchEvent(new CustomEvent(PDF_NATIVE_TEXT_ABANDONED_EVENT, { bubbles: true, detail: Object.freeze({
              target: found?.target ?? null, pageRect: intent.pageRect, text: intent.text, typed: intent.typed,
              x: intent.x, pageNumber: intent.pageNumber,
            }) }));
            intent.typed = "";
          }
        }
        if (intent.typed && process.env.NODE_ENV !== "production")
          markPdfTextLatency("open.typeahead-dropped", { keys: intent.typed.length, live: typeAheadLive(),
            completed: intent.completed, queued: intent.queued, stack: new Error().stack?.split("\n").slice(2, 5).join(" < ").slice(0, 300) });
        document.removeEventListener("keydown", typeAhead, true);
        document.removeEventListener("focusin", deliverOnFocus, true);
        if (laterPress) window.removeEventListener("pointerdown", laterPress, true);
        if (typeAheadFrame) cancelAnimationFrame(typeAheadFrame);
        typeAheadFrame = 0;
        if (stopTypeAheadRef.current === stopTypeAhead) stopTypeAheadRef.current = null;
      };
      const typeAheadLive = () => nativeTextContinuationRef.current === intent ||
        (intent.completed && Date.now() < intent.deliverUntil);
      const typeAhead = (event: KeyboardEvent) => {
        // Given up (a newer press, the line before refusing to let go): the
        // keys kept are still the reader's words for this line.
        if (!typeAheadLive()) { stopTypeAhead("abandon"); return; }
        // The box has taken the keyboard and keys are still kept for it:
        // they go in now, ahead of this key. Left to the next frame, a key
        // typed just after the box was focused went in before them
        // (「…对齐2b01」 for 「b012」); this key then follows them as typed.
        const target = event.target;
        if ((intent.typed || intent.caretMove) && !event.isComposing && target instanceof HTMLTextAreaElement &&
          target.getAttribute(INLINE_INPUT) === "true" && !target.readOnly &&
          (rootRef.current?.contains(target) ?? true)) {
          insertPdfNativeTextTypeAhead(target, intent.typed, intent.caretMove);
          markPdfTextLatency("open.typeahead-delivered", { at: "key", keys: intent.typed.length, move: intent.caretMove });
          intent.typed = "";
          intent.caretMove = null;
          stopTypeAhead();
          return;
        }
        // The line being left keeps the focus while it finishes, its input
        // read-only meanwhile: a key typed there is the next line's, kept
        // here (it went into the read-only input and was lost).
        const leavingInput = event.target instanceof HTMLTextAreaElement &&
          event.target.getAttribute(INLINE_INPUT) === "true" && event.target.readOnly;
        if (event.isComposing || event.ctrlKey || event.metaKey || event.altKey ||
          (isEditable(event.target) && !leavingInput)) return;
        if (event.key === "Escape") { nativeTextContinuationRef.current = null; stopTypeAhead(); return; }
        if (event.key === "Backspace") {
          if (intent.typed) { intent.typed = intent.typed.slice(0, -1); event.preventDefault(); }
          return;
        }
        // Home / End before any key typed for the line: the caret goes
        // there when the box opens (dropped, the words went in where the
        // line was pressed -- into the middle of a word).
        if ((event.key === "End" || event.key === "Home") && !event.shiftKey && !intent.typed && intent.caret) {
          intent.caretMove = event.key === "End" ? "end" : "start";
          event.preventDefault();
          return;
        }
        // A box opening selected as a block takes typing only while
        // 键盘输入触发编辑 is ticked.
        if (event.key.length !== 1 || intent.typed.length >= TYPE_AHEAD_LIMIT ||
          (!intent.caret && !readPdfTextKeyboardInput())) return;
        intent.typed += event.key;
        event.preventDefault();
      };
      // The kept keys go in the moment the box's input has the focus and
      // takes typing: at its focus event, or at the next frame (the focus
      // may be there before this listener is). Left to the frame alone,
      // Safari delivered them 160 ms after the focus, its main thread busy
      
      const deliverTo = (active: Element | null, at: "frame" | "focus") => {
        if (!(active instanceof HTMLTextAreaElement) || active.getAttribute(INLINE_INPUT) !== "true" ||
          active.readOnly || !(rootRef.current?.contains(active) ?? true)) return false;
        if (intent.typed || intent.caretMove) {
          insertPdfNativeTextTypeAhead(active, intent.typed, intent.caretMove);
          markPdfTextLatency("open.typeahead-delivered", { at, keys: intent.typed.length, move: intent.caretMove });
        }
        intent.typed = "";
        intent.caretMove = null;
        stopTypeAhead();
        return true;
      };
      const deliverTypeAhead = () => {
        typeAheadFrame = 0;
        if (!typeAheadLive()) { stopTypeAhead("abandon"); return; }
        if (deliverTo(document.activeElement, "frame")) return;
        typeAheadFrame = requestAnimationFrame(deliverTypeAhead);
      };
      const deliverOnFocus = (event: FocusEvent) => {
        if (!typeAheadLive()) { stopTypeAhead("abandon"); return; }
        const target = event.target;
        if (!(target instanceof HTMLTextAreaElement) || target.getAttribute(INLINE_INPUT) !== "true") return;
        // After the focusing code has put the caret where it wants it (the
        // box focuses and then sets its selection; delivered at the focus
        // event itself, the keys went in at the start of the line). A box
        // still read-only then (its text selected as a block) takes them
        // at the frame, once it takes typing.
        queueMicrotask(() => {
          if (!typeAheadLive()) { stopTypeAhead("abandon"); return; }
          if (document.activeElement === target) deliverTo(target, "focus");
        });
      };
      stopTypeAheadRef.current = stopTypeAhead;
      // A later press anywhere ends this line's keys: a press that names
      // another line keeps its own (a new capture abandons these first);
      // any other press -- on the page, on a box being built again -- left
      // this intent taking every key typed after it, the next line's words
      
      // load: the third line's keys in the second line's intent).
      const createdAt = performance.now();
      laterPress = (event: PointerEvent) => {
        if (event.timeStamp - createdAt < 50 && performance.now() - createdAt < 50) return;
        intent.pressedOver = true;
        // After this press's own handlers: a capture they make replaces the
        // intent (and abandons it) first.
        setTimeout(() => {
          if (nativeTextContinuationRef.current !== intent || intent.completed) return;
          nativeTextContinuationRef.current = null;
          stopTypeAhead("abandon");
        }, 0);
      };
      // On the window: before the press's own handlers, which may let go of the
      // line left, and that settles this intent (below).
      window.addEventListener("pointerdown", laterPress, true);
      document.addEventListener("keydown", typeAhead, true);
      document.addEventListener("focusin", deliverOnFocus, true);
      typeAheadFrame = requestAnimationFrame(deliverTypeAhead);
      // Closing A can change Ribbon height before the browser emits its
      // compatibility mouse events. Those events must not hit a different
      // source box under the old screen point; the captured identity owns it.
      gestureCleanupRef.current?.();
      let releaseTimer: ReturnType<typeof setTimeout> | undefined;
      const cleanup = () => {
        clearTimeout(deadline); clearTimeout(releaseTimer);
        document.removeEventListener("mousedown", consume, true);
        document.removeEventListener("mouseup", consume, true);
        document.removeEventListener("click", consume, true);
        document.removeEventListener("pointerup", pointerUp, true);
        document.removeEventListener("pointercancel", cleanup, true);
        document.removeEventListener("pointerdown", cleanup, true);
        window.removeEventListener("blur", cleanup);
        if (gestureCleanupRef.current === cleanup) gestureCleanupRef.current = null;
      };
      const consume = (event: MouseEvent) => {
        if (event.button !== 0) return;
        event.preventDefault(); event.stopPropagation();
        if (event.type === "click") cleanup();
      };
      const pointerUp = () => { releaseTimer = setTimeout(cleanup, 0); };
      const deadline = setTimeout(cleanup, 30_000);
      gestureCleanupRef.current = cleanup;
      document.addEventListener("mousedown", consume, true);
      document.addEventListener("mouseup", consume, true);
      document.addEventListener("click", consume, true);
      document.addEventListener("pointerup", pointerUp, true);
      document.addEventListener("pointercancel", cleanup, true);
      document.addEventListener("pointerdown", cleanup, true);
      window.addEventListener("blur", cleanup);
      return accepted => {
        if (nativeTextContinuationRef.current !== intent) return;
        // A later press (a click away that let go of the line left, which is
        // what settles this) took the reader elsewhere: the line pressed before
        // it does not open under them. It did, the line just left came back,
        
        if (intent.pressedOver) { nativeTextContinuationRef.current = null; stopTypeAhead("abandon"); return; }
        // The line left did not let go: this line never opens from the
        // intent, and the keys typed for it are written there by the window
        // (PDF_NATIVE_TEXT_ABANDONED_EVENT) -- they were dropped (savesweep
        
        if (!accepted) { nativeTextContinuationRef.current = null; stopTypeAhead("abandon"); return; }
        intent.ready = true;
        intent.readyAt = Date.now();
        intent.deadline = Date.now() + 10_000;
        // The line left has let go. When it staged an edit, the page reads
        // that revision in now, and this line opens queued at once -- in the
        // same update that closes the other (an effect and a frame later, it
        // waited two renders of the window).
        openQueuedNowRef.current(intent);
        setNativeTextContinuationEpoch(epoch => epoch + 1);
      };
    };
    capture.dropAdoption = () => { adoptionRef.current = null; };
    capture.followLine = (target) => {
      const lineId = pdfNativeTextLineIdOf(target);
      const page = target.closest<HTMLElement>(".pdf-viewer__page");
      if (!lineId || !dataItemId || !page) return false;
      const rect = target.getBoundingClientRect();
      adoptionRef.current = {
        dataItemId, pageNumber: Number(page.dataset.pageNumber), sourceRevision: sourceDocumentRevisionRef.current,
        paragraphId: target.getAttribute(PARAGRAPH_ID), itemIndex: target.getAttribute(ITEM_INDEX),
        layoutId: target.getAttribute("data-pdf-native-text-layout-id"), text: target.textContent ?? "",
        lineId, stagedText: target.getAttribute(STAGED_TEXT), x: 0.5, y: 0.5, ready: true,
        deadline: Date.now() + 15_000, inFlight: false, pageRect: pageRelativeRect(target, page),
        workingSessionId: workingSourceRef.current?.sessionId ?? null,
        workingRevision: workingSourceRef.current?.workingRevision ?? null,
        typed: "", caretMove: null, caret: rect.width > 0, completed: true, deliverUntil: 0,
        queued: true, queuedRefused: false,
      };
      setNativeTextContinuationEpoch(epoch => epoch + 1);
      return true;
    };
    capture.readLine = (lineId, pageNumber) => readLineRef.current?.(lineId, pageNumber) ?? null;
    capture.openQueuedNow = () => {
      const intent = nativeTextContinuationRef.current;
      if (intent) openQueuedNowRef.current(intent);
    };
    onCaptureChange?.(capture);
    return () => {
      onCaptureChange?.(null);
      gestureCleanupRef.current?.();
      stopTypeAheadRef.current?.("abandon");
      nativeTextContinuationRef.current = null;
      adoptionRef.current = null;
    };
  }, [currentPage, dataItemId, onCaptureChange, rootRef]);
  useEffect(() => {
    const intent = nativeTextContinuationRef.current;
    const root = rootRef.current;
    if (!intent?.ready || !root) return;
    const cancel = () => {
      if (nativeTextContinuationRef.current === intent) { nativeTextContinuationRef.current = null; stopTypeAheadRef.current?.("abandon"); }
    };
    let frame = 0;
    let disposed = false;
    const attempt = () => {
      frame = 0;
      if (disposed || nativeTextContinuationRef.current !== intent || intent.inFlight) return;
      const revision = sourceDocumentRevisionRef.current;
      const revisionMatches = workingSource
        ? intent.workingSessionId
          ? workingSource.sessionId === intent.workingSessionId && intent.workingRevision !== null &&
            workingSource.workingRevision >= intent.workingRevision && workingSource.workingRevision <= intent.workingRevision + 1
          : workingSource.originDocumentRevision === intent.sourceRevision && workingSource.workingRevision === 1
        : typeof revision === "number" && typeof intent.sourceRevision === "number"
        ? revision >= intent.sourceRevision && revision <= intent.sourceRevision + 1
        : revision === intent.sourceRevision;
      if (Date.now() >= intent.deadline || dataItemId !== intent.dataItemId || currentPage !== intent.pageNumber ||
        !modeActive || !writable) { cancel(); return; }
      // A revision further on than the one the press was made on (the line
      // left was saved another way meanwhile: a text box written straight
      // to the file, the page read in again) no longer drops the press and
      
      // the line opens there when the box at the press is the same line,
      // its words unchanged.
      const sameWords = (span: HTMLElement) => {
        const words = (span.textContent || "").replace(/\s+/g, " ").trim();
        return words === intent.text.replace(/\s+/g, " ").trim() ||
          (intent.stagedText !== null && words === intent.stagedText.replace(/\s+/g, " ").trim());
      };
      // The page's boxes held for a revision being read in (an edit the
      // working copy just staged): the line opens at once on the page as
      // shown -- WPS switches lines instantly -- and is opened on the new
      // revision once its boxes are back, taking that box over.
      const opening = queuedRef.current;
      const queuedOpen = Boolean(opening) && !interactive && !editingActive && !intent.queuedRefused &&
        adoptionRef.current === null;
      if (!queuedOpen && (editingActive || !interactive)) return;
      const found = findPdfNativeTextIntentTarget(root, intent, queuedOpen);
      if (!found) return;
      // Another working copy's lineage (a session this press was not made
      // in) is never the line pressed, whatever its words.
      const foreignLineage = Boolean(workingSource && intent.workingSessionId &&
        workingSource.sessionId !== intent.workingSessionId);
      if (!revisionMatches && (queuedOpen || foreignLineage || !sameWords(found.target))) return;
      const { target, point } = found;
      // Normal activation rebinds the current bytes, font, layout and authority.
      intent.inFlight = true;
      if (queuedOpen) {
        if (!opening!.open(target, point, !intent.caret, status => {
          intent.inFlight = false;
          if (status === "queued") { queuedOpened(intent, target); return; }
          // Not openable queued (a box already written as a text box, a
          // line without its layout): it waits for its boxes as before,
          // and is not asked to open queued again.
          intent.queuedRefused = true;
          if (nativeTextContinuationRef.current === intent) setNativeTextContinuationEpoch(epoch => epoch + 1);
        })) { intent.inFlight = false; intent.queuedRefused = true; }
        return;
      }
      if (!activate(target, point, false, status => {
        intent.inFlight = false;
        if (status === "accepted") {
          // Spent, except for the keys typed ahead: they still wait for
          // the opened box's input to take focus.
          intent.completed = true;
          intent.deliverUntil = Date.now() + 3_000;
          if (nativeTextContinuationRef.current === intent) nativeTextContinuationRef.current = null;
        } else if (status !== "transient-refusal") cancel();
        else if (nativeTextContinuationRef.current === intent) setNativeTextContinuationEpoch(epoch => epoch + 1);
        // The press that left the line before selects this one as a block
        // (WPS); a double click puts its caret in.
      }, !intent.caret)) intent.inFlight = false;
    };
    const schedule = () => { if (!frame) frame = requestAnimationFrame(attempt); };
    const observer = new MutationObserver(schedule);
    observer.observe(root, { childList: true, subtree: true, attributes: true,
      attributeFilter: [EDITABLE, EDIT_BOX, PARAGRAPH_ID, ITEM_INDEX, "data-pdf-native-text-layout-id"] });
    // The page's boxes held for the reader's own edit going up (a Delete's
    // steps, a revision being read in): the press waits for them while they
    // are, not 10 s. It ran out while a slow save kept them held, and the
    
    // pressed while a Delete was saved, 4 s a step). A newer press still
    // takes its place at once (below).
    const waitCap = (intent.readyAt ?? Date.now()) + PRESS_WAIT_FOR_HELD_PAGE_MS;
    if (!interactive) intent.deadline = Math.max(intent.deadline, Math.min(Date.now() + 10_000, waitCap));
    let timer: ReturnType<typeof setTimeout> | 0 = 0;
    const expire = () => {
      if (!disposed && !interactive && Date.now() < waitCap && nativeTextContinuationRef.current === intent) {
        intent.deadline = Math.min(Date.now() + 10_000, waitCap);
        timer = setTimeout(expire, Math.max(0, intent.deadline - Date.now()));
        return;
      }
      cancel();
    };
    timer = setTimeout(expire, Math.max(0, intent.deadline - Date.now()));
    // A newer real gesture supersedes this old intent, including a blank click.
    document.addEventListener("pointerdown", cancel, true);
    schedule();
    return () => { disposed = true; observer.disconnect(); if (timer) clearTimeout(timer); if (frame) cancelAnimationFrame(frame); document.removeEventListener("pointerdown", cancel, true); };
  }, [nativeTextContinuationEpoch, currentPage, dataItemId, modeActive,
    writable, editingActive, interactive, activate, rootRef, sourceRevision, workingSource]);
  // The line opened queued is opened on the revision the moment its boxes are
  // back, taking that box over (its draft, caret and any finish the reader asked
  // for by leaving it). Its own box is open meanwhile: this opening runs while
  // editing is active, by design. In the commit that gives the boxes back, not
  // a frame or a DOM change later: its revision's text went up with them, and
  // the line is found by its id at once.
  useLayoutEffect(() => {
    const intent = adoptionRef.current;
    const root = rootRef.current;
    if (!intent || !root) return;
    const drop = () => { if (adoptionRef.current === intent) adoptionRef.current = null; };
    let frame = 0;
    let disposed = false;
    const attempt = () => {
      frame = 0;
      if (disposed || adoptionRef.current !== intent || intent.inFlight) return;
      if (Date.now() >= intent.deadline || dataItemId !== intent.dataItemId || currentPage !== intent.pageNumber ||
        !modeActive || !writable) { drop(); return; }
      if (!interactive) return;
      const found = findPdfNativeTextIntentTarget(root, intent, false, true, !intent.lookedFor);
      intent.lookedFor = true;
      if (!found) return;
      // Its line is being moved and the page still reads it where it started
      // (the move's revision is not up): the box opened on it stands where
      // the reader put it, and is taken over once the line is read there.
      // Taken over here, it went back to where the line had been.
      if (pdfNativeTextLineIsGoing(found.target)) return;
      intent.inFlight = true;
      if (!activate(found.target, found.point, false, status => {
        intent.inFlight = false;
        if (adoptionRef.current !== intent) return;
        if (status === "transient-refusal") { setNativeTextContinuationEpoch(epoch => epoch + 1); return; }
        drop();
      }, !intent.caret, "adopt")) intent.inFlight = false;
    };
    const schedule = () => { if (!frame) frame = requestAnimationFrame(attempt); };
    const observer = new MutationObserver(schedule);
    observer.observe(root, { childList: true, subtree: true, attributes: true,
      attributeFilter: [EDITABLE, EDIT_BOX, PARAGRAPH_ID, ITEM_INDEX, "data-pdf-native-text-layout-id"] });
    const timer = setTimeout(() => { drop(); }, Math.max(0, intent.deadline - Date.now()));
    attempt();
    return () => { disposed = true; observer.disconnect(); clearTimeout(timer); if (frame) cancelAnimationFrame(frame); };
  }, [nativeTextContinuationEpoch, currentPage, dataItemId, modeActive, writable, interactive, activate, rootRef]);

}
