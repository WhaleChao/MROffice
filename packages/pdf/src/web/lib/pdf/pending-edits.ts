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
  for (let attempt = 0; attempt < ATTEMPTS; attempt++) {
    try {
      await apiPost(`/api/v1/pages/${encodeURIComponent(input.pageId)}/pdf-pending-edits`, {
        edit_id: input.editId, base_sha256: baseSha256, sequence: input.sequence, payload: input.payload,
      });
      note(input.pageId, input.editId);
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

