/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { apiGet, apiPost } from "@mrnote/core/web/lib/api.ts";
import { markPdfTextLatency } from "./text-edit-latency";

export type PdfPendingEditPayload = Readonly<Record<string, unknown>>;

const EDIT_ID = /^[A-Za-z0-9][A-Za-z0-9:_-]{0,63}$/u;
const SHA256 = /^[0-9a-f]{64}$/u;
const ATTEMPTS = 3;

/** Edits this window has saved as rows, by page: the ids the server answered for. */
const saved = new Map<string, Set<string>>();
/** What each edit was saved as, so a Redo saves it again (by page and id). */
const inputs = new Map<string, Parameters<typeof savePdfPendingEdit>[0]>();
const key = (pageId: string, editId: string) => `${pageId}\u0000${editId}`;
const listeners = new Set<() => void>();
/** How many times a page's window has closed (forgetPdfPendingEditsSavedHere):
 * a save still on its way then is not the next window's own. */
const closings = new Map<string, number>();

function note(pageId: string, editId: string): void {
  let ids = saved.get(pageId);
  if (!ids) { ids = new Set(); saved.set(pageId, ids); }
  ids.add(editId);
  for (const listener of listeners) {
    try { listener(); } catch { /* a listener cannot stop another */ }
  }
}

/** Whether this window has had `editId` saved as a row of `pageId`. */
export function isPdfPendingEditSaved(pageId: string, editId: string): boolean {
  return saved.get(pageId)?.has(editId) ?? false;
}

/** The kinds of edit saved with the engine commands that write them. */
const RECIPE_KINDS: ReadonlySet<unknown> = new Set(["format", "move", "box-format"]);

/** Whether a row holds what writes its edit without the window that saved
 * it: a plain retyping (its selection and words), or a format, a box's
 * paragraph format or a move saved with its engine commands (`attempts`)
 * whose fonts the server can read itself. */
export function pdfPendingEditPayloadIsWritable(payload: PdfPendingEditPayload | null | undefined): boolean {
  if (!payload) return false;
  if (payload.kind === "text") return true;
  return RECIPE_KINDS.has(payload.kind) && Array.isArray(payload.attempts) && payload.attempts.length > 0 &&
    payload.serverWritable !== false;
}

/** Whether `editId` is saved as a row a window opened later -- or the
 * server, when this one goes away -- writes on its own
 * (pdfPendingEditPayloadIsWritable). Any other edit is saved only with its
 * file version. */
export function isPdfPendingEditDurable(pageId: string, editId: string): boolean {
  return isPdfPendingEditSaved(pageId, editId) && pdfPendingEditPayloadIsWritable(inputs.get(key(pageId, editId))?.payload);
}

/** Called whenever an edit is saved as a row. */
export function onPdfPendingEditSaved(listener: () => void): () => void {
  listeners.add(listener);
  return () => { listeners.delete(listener); };
}

/**
 * Save one edit as a row. Resolves true once the server holds it (a retry of
 * the same edit answers the row it made), false when it could not be saved
 * that way -- the edit is then saved only with its file version, as before.
 * Never throws.
 */
export async function savePdfPendingEdit(input: Readonly<{
  pageId: string; editId: string; baseSha256: string; sequence: number; payload: PdfPendingEditPayload;
}>): Promise<boolean> {
  const baseSha256 = input.baseSha256.toLowerCase();
  if (!input.pageId || !EDIT_ID.test(input.editId) || !SHA256.test(baseSha256) ||
    !Number.isInteger(input.sequence) || input.sequence < 0) return false;
  inputs.set(key(input.pageId, input.editId), input);
  while (inputs.size > 500) inputs.delete(inputs.keys().next().value as string);
  const posted = postPdfPendingEdit(input, baseSha256);
  notePdfPendingEditSaving(input.pageId, input.editId, posted);
  return posted;
}

async function postPdfPendingEdit(input: Parameters<typeof savePdfPendingEdit>[0], baseSha256: string): Promise<boolean> {
  const startedAt = performance.now();
  const closing = closings.get(input.pageId) ?? 0;
  for (let attempt = 0; attempt < ATTEMPTS; attempt++) {
    try {
      await apiPost(`/api/v1/pages/${encodeURIComponent(input.pageId)}/pdf-pending-edits`, {
        edit_id: input.editId, base_sha256: baseSha256, sequence: input.sequence, payload: input.payload,
      });
      // (Its window closed while it was on its way: the row is saved, and
      // is the next window's to write like any other it finds.)
      if ((closings.get(input.pageId) ?? 0) === closing) note(input.pageId, input.editId);
      markPdfTextLatency("save.pending-saved", { ms: Math.round(performance.now() - startedAt), attempt });
      return true;
    } catch (error) {
      const status = (error as { status?: unknown })?.status;
      // A refusal of the edit itself (not writable, malformed, too large):
      // asking again gets the same.
      if (typeof status === "number" && status >= 400 && status < 500 && status !== 408 && status !== 429) break;
      await new Promise(resolve => { setTimeout(resolve, 250 * 2 ** attempt); });
    }
  }
  markPdfTextLatency("save.pending-unsaved", { ms: Math.round(performance.now() - startedAt) });
  return false;
}

/** Saves of a row still on their way (an edit's fonts fetched before it is
 * posted), by page and id: a withdrawal waits for them, so a row is never
 * withdrawn before it is made -- the save arriving after would leave it
 * pending, and the server would write an edit taken back. */
const saving = new Map<string, Promise<unknown>>();

/** Note a save of `editId` on its way before savePdfPendingEdit is called
 * (what it saves is still being gathered). */
export function notePdfPendingEditSaving(pageId: string, editId: string, save: Promise<unknown>): void {
  const id = key(pageId, editId);
  const settled = save.catch(() => undefined).finally(() => {
    if (saving.get(id) === settled) saving.delete(id);
  });
  saving.set(id, settled);
}

/** The rows of a page this window saved are no longer its own to answer for
 * (the window closed): a window opened on the page later reads every row no
 * file version carries, these among them. */
export function forgetPdfPendingEditsSavedHere(pageId: string): void {
  saved.delete(pageId);
  closings.set(pageId, (closings.get(pageId) ?? 0) + 1);
}

/** Take an edit saved as a row back (Undo before it was written). Never throws. */
export async function withdrawPdfPendingEdit(pageId: string, editId: string): Promise<void> {
  if (!pageId || !EDIT_ID.test(editId)) return;
  await saving.get(key(pageId, editId));
  try {
    await apiPost(`/api/v1/pages/${encodeURIComponent(pageId)}/pdf-pending-edits/${encodeURIComponent(editId)}/withdraw`, {});
  } catch { /* an edit not withdrawn is written with its version, as any */ }
  saved.get(pageId)?.delete(editId);
}

/** Save again an edit taken back and put back (Redo before it was written):
 * the server holds its withdrawn row pending again. Never throws. */
export async function resavePdfPendingEdit(pageId: string, editId: string): Promise<boolean> {
  const input = inputs.get(key(pageId, editId));
  return input ? savePdfPendingEdit(input) : false;
}

/** Test seam. */
export function resetPdfPendingEditsForTests(): void {
  keptTextBlocksTaken = null;
  saved.clear();
  closings.clear();
  inputs.clear();
  saving.clear();
  listeners.clear();
}

let keptTextBlocksTaken: Promise<boolean> | null = null;

/**
 * Whether the API this window reaches takes the text blocks an edit set on a
 * revision published as its bytes (NativeBytesPublishRecord.keptTextBlocks,
 * native-page-box-map.ts). Asked once a window: an API that predates them
 * answers no route, and the blocks stay with the open page alone. Never
 * throws.
 */
export function pdfNativeJournalTakesKeptTextBlocks(pageId: string): Promise<boolean> {
  if (!pageId) return Promise.resolve(false);
  keptTextBlocksTaken ??= apiGet<{ kept_text_blocks?: unknown }>(
    `/api/v1/pages/${encodeURIComponent(pageId)}/pdf-native-capabilities`)
    .then((answer) => answer?.kept_text_blocks === true)
    .catch((error) => {
      // An API without the route says no for good; a failure to reach it is
      // asked again with the next edit.
      const status = (error as { status?: unknown })?.status;
      if (status !== 404 && status !== 405) keptTextBlocksTaken = null;
      return false;
    });
  return keptTextBlocksTaken;
}

let chainSaveTaken: Promise<boolean> | null = null;

/**
 * Whether the API this window reaches saves a chain of queued revisions in
 * one request (lib/pdf/native-chain-save). Asked once a window, as above: an
 * API that predates it says no, and the queue goes one save at a time.
 * Never throws.
 */
export function pdfNativeChainSaveTaken(pageId: string): Promise<boolean> {
  if (!pageId) return Promise.resolve(false);
  chainSaveTaken ??= apiGet<{ chain_save?: unknown }>(
    `/api/v1/pages/${encodeURIComponent(pageId)}/pdf-native-capabilities`)
    .then((answer) => answer?.chain_save === true)
    .catch((error) => {
      const status = (error as { status?: unknown })?.status;
      if (status !== 404 && status !== 405) chainSaveTaken = null;
      return false;
    });
  return chainSaveTaken;
}


/** A row no file version of the page carries yet, as the API lists it. */
export type PdfPendingEditRow = Readonly<{
  editId: string;
  /** `pending`: the server's writer has not been to it (or will be again).
   * `failed`: it could not write it on the version it was made on. Either
   * way it is this window's to write on the version the page has now. */
  status: "pending" | "failed";
  sequence: number;
  /** Why the server's writer could not (its code), when it said. */
  failure: string | null;
  /** Whether the reader of this window made the edit; null when the API
   * does not say (one that predates the question: nothing is reported). */
  mine: boolean | null;
  /** How long ago it was saved, by the server's clock; null when the API
   * does not say. */
  ageSeconds: number | null;
  payload: PdfPendingEditPayload | null;
}>;

/**
 * The page's rows no file version carries: null when they could not be read
 * (asked again later), an empty list when there are none -- or none this
 * reader may read (a visitor, a page they cannot see). Never throws.
 */
export async function listPdfPendingEditRows(pageId: string): Promise<readonly PdfPendingEditRow[] | null> {
  if (!pageId) return null;
  let answer: unknown;
  try {
    answer = await apiGet<unknown>(`/api/v1/pages/${encodeURIComponent(pageId)}/pdf-pending-edits`);
  } catch (error) {
    const status = (error as { status?: unknown })?.status;
    // The reader's own refusal answers the same however often it is asked.
    return typeof status === "number" && status >= 400 && status < 500 && status !== 408 && status !== 429 ? [] : null;
  }
  // (Anything but a list is an API that keeps no rows.)
  if (!Array.isArray(answer)) return [];
  const rows: PdfPendingEditRow[] = [];
  for (const entry of answer as readonly Record<string, unknown>[]) {
    if (!entry || typeof entry !== "object" || typeof entry.edit_id !== "string" || !EDIT_ID.test(entry.edit_id) ||
      (entry.status !== "pending" && entry.status !== "failed")) continue;
    const payload = entry.payload && typeof entry.payload === "object" && !Array.isArray(entry.payload)
      ? entry.payload as PdfPendingEditPayload : null;
    rows.push(Object.freeze({ editId: entry.edit_id, status: entry.status,
      sequence: typeof entry.sequence === "number" ? entry.sequence : 0,
      failure: typeof entry.failure === "string" && entry.failure ? entry.failure : null,
      mine: typeof entry.mine === "boolean" ? entry.mine : null,
      ageSeconds: typeof entry.age_seconds === "number" && Number.isFinite(entry.age_seconds) ? entry.age_seconds : null,
      payload }));
  }
  return rows;
}

/** Settle a row its maker's window could not write and has told them of
 * (`reported`): it is offered to no window again. False when the server did
 * not take it (the row is then told of again when the page is next opened).
 * Never throws. */
export async function reportPdfPendingEditUnwritten(pageId: string, editId: string, reason: string): Promise<boolean> {
  if (!pageId || !EDIT_ID.test(editId)) return false;
  try {
    await apiPost(`/api/v1/pages/${encodeURIComponent(pageId)}/pdf-pending-edits/${encodeURIComponent(editId)}/unwritten`,
      { reason: reason.replace(/[^A-Za-z0-9 _:.,/()-]/gu, "?").slice(0, 200) });
    return true;
  } catch {
    return false;
  }
}

/** Whether a window can try to write a row at all: its line and its words,
 * and for a format, a paragraph format or a move the engine commands saved
 * with it. (A row the server cannot write for its fonts -- `serverWritable:
 * false` -- is still a window's to try: it reads the faces itself.) */
export function pdfPendingEditRowIsWindowWritable(row: PdfPendingEditRow): boolean {
  const payload = row.payload;
  const selection = payload?.selection as { pageNumber?: unknown; item?: unknown } | null | undefined;
  if (!payload || typeof payload.replacementText !== "string" || !selection || typeof selection !== "object" ||
    !Number.isInteger(selection.pageNumber) || !selection.item || typeof selection.item !== "object") return false;
  return payload.kind === "text" ||
    (RECIPE_KINDS.has(payload.kind) && Array.isArray(payload.attempts) && payload.attempts.length > 0);
}

/** Whether a row carries words the reader typed (a retyping, or a format
 * left with the line's words changed): kept as the line's draft when the
 * row cannot be written. */
export function pdfPendingEditRowHasTypedWords(row: PdfPendingEditRow): boolean {
  return typeof row.payload?.replacementText === "string" &&
    (row.payload.kind === "text" || row.payload.words === true);
}

/** The reason said for a line the page no longer shows as it read when the
 * edit was made (the window's own verdict; the notice words it). */
export const PDF_PENDING_ROW_PAGE_CHANGED = "page_changed";
/** The reason for a row saved without what writes it. */
export const PDF_PENDING_ROW_UNWRITABLE = "unwritable";

/**
 * What became of one try at writing a row on the page as it stands:
 * `staged` (it is in this window's queue, saved as any edit of its own);
 * `later` (nothing was decided: the planner or the page was out of reach,
 * the queue was busy -- tried again); or refused, with why (the line is not
 * there as it read, the engine will not write it).
 */
export type PdfPendingRowAttempt = "staged" | "later" | Readonly<{ refused: string }>;

/** Failures that decide nothing about the edit: a planner, a Worker or the
 * engine host out of reach, a request cancelled. By name: the classes live
 * with the callers (the working session, the engine client). */
const UNDECIDED_ERROR_NAMES: ReadonlySet<string> = new Set(["PdfWorkingPlannerUnavailableError",
  "PdfNativeBrowserPlanUnavailable", "PdfNativeBrowserPrepareUnavailable", "PdfNativeTextWorkerUnavailable",
  "EngineUnavailableError", "AbortError", "TimeoutError"]);
/** ...and the working copy's own "not now": it moved under the try, another
 * edit is being prepared, it waits for a save, it closed. */
const UNDECIDED_ERROR_MESSAGES = new RegExp("PDF working document is closed|PDF working session authority|" +
  "PDF working display|changed during preparation|stale or already being prepared|" +
  "no longer owns the working source|must finish saving|memory budget", "u");

/** What a failed try at staging a row says of the row: `later` when nothing
 * was decided about the edit, a refusal (the error's own words, cut short)
 * otherwise. */
export function pdfPendingRowAttemptOfError(error: unknown): PdfPendingRowAttempt {
  const name = (error as { name?: unknown } | null)?.name;
  const message = String((error as { message?: unknown } | null)?.message ?? error ?? "");
  if ((typeof name === "string" && UNDECIDED_ERROR_NAMES.has(name)) || UNDECIDED_ERROR_MESSAGES.test(message)) return "later";
  return Object.freeze({ refused: message.slice(0, 120) || "refused" });
}


















export async function stagePdfPendingRowWords(input: Readonly<{
  compile: () => Promise<void>;
  /** The copy holds the edit although the compile threw. */
  taken: () => boolean;
  byEngine: (compileRefusal: unknown) => Promise<PdfPendingRowAttempt>;
}>): Promise<PdfPendingRowAttempt> {
  let compileRefusal: unknown;
  try {
    await input.compile();
    return "staged";
  } catch (error) {
    if (input.taken()) return "staged";
    if (pdfPendingRowAttemptOfError(error) === "later") throw error;
    markPdfTextLatency("working.server-pending-engine", {
      why: String((error as { message?: unknown } | null)?.message ?? error).slice(0, 80) });
    compileRefusal = error;
  }
  return input.byEngine(compileRefusal);
}

/** What a window remembers of the rows it tried, across its tries. */
export type PdfPendingRowsLedger = {
  /** Refusals so far, by edit: a row is given up only when refused twice. */
  readonly refusals: Map<string, number>;
  /** Rows this window has told its reader of (not said twice). */
  readonly told: Set<string>;
};

export function createPdfPendingRowsLedger(): PdfPendingRowsLedger {
  return { refusals: new Map(), told: new Set() };
}

export type PdfPendingRowUnwritten = Readonly<{
  row: PdfPendingEditRow;
  reason: string;
  /** The reader's words are kept as the line's draft. */
  draftKept: boolean;
  /** The server took the report (false: told again at the next open). */
  settled: boolean;
}>;

export type PdfPendingRowsOutcome = Readonly<{
  /** Rows now in this window's queue. */
  staged: number;
  /** Rows of this reader that can be written nowhere: to be said now. */
  unwritten: readonly PdfPendingRowUnwritten[];
  /** Rows nothing was decided about: this window tries them again. */
  again: number;
  /** Of those, this reader's own (said, if this window never gets to them). */
  againMine: number;
  /** Another reader's rows this window could not write: theirs to be told of. */
  left: number;
}>;

/** A refusal is taken for the row's last word when the same window meets it
 * this many times: once may be the page still settling under it. */
export const PDF_PENDING_ROW_REFUSALS_TO_GIVE_UP = 2;
/** A row still `pending` and younger than this may be its own window's to
 * write yet (another tab, seconds into its save): it is not given up on --
 * nor its reader told it was not written -- before the server's own writer
 * would take it for abandoned (MATERIALIZE_AFTER_SECONDS). */
export const PDF_PENDING_ROW_ABANDONED_AFTER_SECONDS = 120;

/** Whether nobody else is about to write the row: the server's writer has
 * said it could not, or its window has had its time. */
function rowIsNobodysElse(row: PdfPendingEditRow): boolean {
  return row.status === "failed" ||
    (row.ageSeconds !== null && row.ageSeconds >= PDF_PENDING_ROW_ABANDONED_AFTER_SECONDS);
}

/**
 * Write the rows no file version carries, as this window's own edits, on the
 * version the page has now -- and account for every one of them. Each row
 * ends this call staged, to be tried again (`again`), told to the reader who
 * made it (`unwritten`: refused twice here, or saved without what writes it,
 * and nobody else's to write any more; their words kept as a draft and the
 * row settled `reported`), or left for the reader who made it (`left`).
 * None is dropped without one of these.
 *
 * `isHandled`: rows that are not this call's -- the page's journal carries
 * them, or this window holds them itself (its queue, its recovery, an edit
 * it is still staging). `canWrite`: whether this window can write the page
 * now. `stage`: one try at one row (a throw is read by
 * pdfPendingRowAttemptOfError). Never throws.
 */
export async function writePdfPendingEditRows(input: Readonly<{
  pageId: string;
  rows: readonly PdfPendingEditRow[];
  ledger: PdfPendingRowsLedger;
  isHandled: (editId: string) => boolean;
  canWrite: () => boolean | Promise<boolean>;
  stage: (row: PdfPendingEditRow) => Promise<PdfPendingRowAttempt>;
  keepDraft: (row: PdfPendingEditRow) => boolean;
  report?: (pageId: string, editId: string, reason: string) => Promise<boolean>;
}>): Promise<PdfPendingRowsOutcome> {
  const report = input.report ?? reportPdfPendingEditUnwritten;
  const { ledger } = input;
  const open = input.rows.filter((row) => !ledger.told.has(row.editId) && !input.isHandled(row.editId));
  let staged = 0, again = 0, againMine = 0, left = 0;
  const unwritten: PdfPendingRowUnwritten[] = [];
  if (!open.length) return Object.freeze({ staged, unwritten, again, againMine, left });
  let writable = false;
  try { writable = await input.canWrite(); } catch { writable = false; }
  for (const row of open) {
    let attempt: PdfPendingRowAttempt;
    const writableHere = pdfPendingEditRowIsWindowWritable(row);
    if (!writableHere) {
      attempt = { refused: row.failure ?? PDF_PENDING_ROW_UNWRITABLE };
    } else if (!writable || (row.payload?.serverWritable === false && !rowIsNobodysElse(row))) {
      // (A format the server cannot write for its fonts is written by its
      // own window with its file version, and was never read as saved
      // before that: while that window may be at it, it is left to it.)
      attempt = "later";
    } else if (input.isHandled(row.editId)) {
      // Taken up by this window itself while the rows before it were tried.
      continue;
    } else {
      try { attempt = await input.stage(row); }
      catch (error) { attempt = pdfPendingRowAttemptOfError(error); }
    }
    if (attempt === "staged") {
      staged += 1;
      ledger.refusals.delete(row.editId);
      continue;
    }
    if (attempt === "later") {
      again += 1;
      if (row.mine === true) againMine += 1;
      continue;
    }
    const refusals = (ledger.refusals.get(row.editId) ?? 0) + 1;
    ledger.refusals.set(row.editId, refusals);
    if ((writableHere && refusals < PDF_PENDING_ROW_REFUSALS_TO_GIVE_UP) || !rowIsNobodysElse(row)) {
      again += 1;
      if (row.mine === true) againMine += 1;
      continue;
    }
    if (row.mine !== true) {
      // Another reader's edit (or an API that does not say whose): theirs
      // to be told of, in their window. It stays as it is.
      left += 1;
      continue;
    }
    let draftKept = false;
    if (pdfPendingEditRowHasTypedWords(row)) {
      try { draftKept = input.keepDraft(row); } catch { draftKept = false; }
    }
    ledger.told.add(row.editId);
    const settled = await report(input.pageId, row.editId, attempt.refused).catch(() => false);
    unwritten.push(Object.freeze({ row, reason: attempt.refused, draftKept, settled }));
  }
  return Object.freeze({ staged, unwritten, again, againMine, left });
}
