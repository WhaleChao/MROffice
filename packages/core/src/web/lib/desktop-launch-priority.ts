/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { IS_DESKTOP_CLASSIC } from "./desktop-classic";
import { documentShown, pageSpeaks } from "./desktop-instant";
import { PRESENTATION_FIRST_SLIDE_PAINTED_EVENT } from "./presentation/first-slide-painted";

/** Reads of the first screen no document's opening waits for. */
const DEFERRED_UNTIL_IMPORT =
  /^\/api\/v1\/(office-data\/|billing\/|notifications\/|desktop\/sync\/status|collaboration-windows\/instances\/|messaging\/|notebook-shelf\/opened)/;
/** A document's import request (a Word document, a workbook, a PDF: the
 * batch's PUT; a deck: its upload's reservation). */
const LAUNCH_IMPORT = /^\/api\/v1\/notebooks\/[^/?]+\/(import-batches\/[^/?]+|source-uploads)$/;
export const LONGEST_HOLD_MS = 2_500;

type Hold = { promise: Promise<void>; release: () => void };
let hold: Hold | null = null;
/** This page load's hold was made and let go: never made again. */
let done = false;

function launchOpensDocuments(): boolean {
  try {
    return new URL(window.location.href).searchParams.get("openFiles") === "desktop";
  } catch {
    return false;
  }
}

function currentHold(): Hold | null {
  if (hold || done || !IS_DESKTOP_CLASSIC || typeof window === "undefined") return hold;
  if (!launchOpensDocuments()) return null;
  let release: () => void = () => undefined;
  const promise = new Promise<void>((resolve) => {
    release = resolve;
  });
  const made: Hold = { promise, release: () => undefined };
  const timer = window.setTimeout(() => made.release(), LONGEST_HOLD_MS);
  made.release = () => {
    window.clearTimeout(timer);
    if (hold === made) hold = null;
    done = true;
    release();
  };
  hold = made;
  return made;
}

/** The page has nothing of its launch to import (every document already a
 * page, or none): the held reads go now. */
export function releaseLaunchImportHold(): void {
  if (hold) hold.release();
  else done = true;
}














const DEFERRED_WHILE_OPENING =
  /^\/api\/v1\/(notebooks|notebooks\/[^/?]+\/(pages|folders|study-assets)|pages\/[^/?]+\/(macros|mention-candidates|changes-since-read|native-xlsx-copy-capabilities|spreadsheet\/data-from-picture\/capability)|data-items\/[^/?]+\/(annotations|preview)|attachments\/page\/[^/?]+\/transcription|editor-plugins|resources(\/[^?]*)?|ai\/turn-runs\/copilot\/pending|collaboration-windows\/instances\/accessible|auth\/me|workspaces\/me\/storage-usage)(\?|$)/;
/** The page catalogue's read is a POST (lib/notebook-page-catalogue). */
const DEFERRED_POST_WHILE_OPENING = /^\/api\/v1\/notebooks\/[^/?]+\/page-changes(\?|$)/;




export const OPENING_HOLD_MS = 4_000;


export const OPENING_IMPORT_HOLD_MS = 60_000;
const OPENING_CHECK_MS = 50;
/** The event the notebook's page imports opened documents on (lib/open-files/desktop
 * DESKTOP_OPEN_FILES_EVENT; named here: that module reads the API through this one). */
const OPEN_FILES_EVENT = "mrnote:desktop-open-files";

let opening: Hold | null = null;

function openingHoldOff(): boolean {
  return (window as { __openingHoldOff?: boolean }).__openingHoldOff === true;
}

/** The document the current hold waits for (null: any shown in the front window). */
let openingName: string | null = null;

/** When the current hold lets go at the latest. */
let openingDeadline = 0;

/** Hold the background reads while the document named `name` opens (see
 * DEFERRED_WHILE_OPENING), for at most `longestMs`. A hold already on waits
 * for this one instead, no shorter than it was to: what it holds is not let
 * go meanwhile. */
export function holdReadsWhileDocumentOpens(name: string | null, longestMs = OPENING_HOLD_MS): void {
  if (!IS_DESKTOP_CLASSIC || typeof window === "undefined" || openingHoldOff()) return;
  openingName = name;
  openingDeadline = Math.max(opening ? openingDeadline : 0, Date.now() + longestMs);
  if (opening) return;
  let release: () => void = () => undefined;
  const promise = new Promise<void>((resolve) => {
    release = resolve;
  });
  const made: Hold = { promise, release: () => undefined };
  const painted = () => made.release();
  const timer = window.setInterval(() => {
    if (Date.now() >= openingDeadline || pageSpeaks() || documentShown(document, openingName)) made.release();
  }, OPENING_CHECK_MS);
  window.addEventListener(PRESENTATION_FIRST_SLIDE_PAINTED_EVENT, painted);
  made.release = () => {
    window.clearInterval(timer);
    window.removeEventListener(PRESENTATION_FIRST_SLIDE_PAINTED_EVENT, painted);
    if (opening === made) opening = null;
    release();
  };
  opening = made;
}

if (IS_DESKTOP_CLASSIC && typeof window !== "undefined") {
  window.addEventListener(OPEN_FILES_EVENT, (event) => {
    const files = (event as CustomEvent<unknown>).detail;
    const first = Array.isArray(files) ? files[0] : null;
    if (!first) return;
    // Not for a PDF: its page counts as shown (canvas and text layer in the
    // page) before it is drawn, and the reads let go then were answered on
    
    const named = typeof (first as { name?: unknown }).name === "string" ? (first as File).name : "";
    if (/\.pdf$/i.test(named) || (first as { type?: unknown }).type === "application/pdf") return;
    holdReadsWhileDocumentOpens(typeof (first as { name?: unknown }).name === "string" ? (first as File).name : null, OPENING_IMPORT_HOLD_MS);
  }, true);
}

/** The import of a document opened from the system has answered (its page
 * made, or the import failed): the hold waits OPENING_HOLD_MS more at most. */
export function noteOpenedDocumentImported(): void {
  if (!opening) return;
  openingDeadline = Date.now() + OPENING_HOLD_MS;
}

/**
 * Settles once no document opened from the system is opening (at once when
 * none is): work of the notebook's own that is no request of the API's
 * client but competes for the same API and main thread -- a page list's
 * thumbnails, which read each new page whole and draw a PDF's first page --
 * waits for it as the reads above do.
 */
export function whileDocumentOpens(): Promise<void> {
  return opening ? opening.promise : Promise.resolve();
}

function openingWait(path: string, method: string): Promise<void> | null {
  if (!opening) return null;
  const verb = method.toUpperCase();
  if (verb === "GET" ? DEFERRED_WHILE_OPENING.test(path) : verb === "POST" && DEFERRED_POST_WHILE_OPENING.test(path)) {
    return opening.promise;
  }
  return null;
}

/**
 * Before an API request goes out: what it waits for first (a held read), or
 * null. A launch's import lets the held reads go once it is itself handed to
 * the network (the next task: its fetch is queued for a connection first).
 */
export function launchPriorityWait(path: string, method: string): Promise<void> | null {
  const whileOpening = openingWait(path, method);
  if (whileOpening) return whileOpening;
  const current = currentHold();
  if (!current) return null;
  const bare = path.split("?")[0];
  const verb = method.toUpperCase();
  if ((verb === "PUT" || verb === "POST") && LAUNCH_IMPORT.test(bare)) {
    window.setTimeout(() => current.release(), 0);
    return null;
  }
  return DEFERRED_UNTIL_IMPORT.test(path) ? current.promise : null;
}

/** Tests: a fresh page load. */
export function forgetLaunchPriorityForTests(): void {
  hold?.release();
  hold = null;
  done = false;
  opening?.release();
  opening = null;
}
