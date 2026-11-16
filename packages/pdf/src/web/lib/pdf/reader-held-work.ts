/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** A write is running, or the last try failed and another is coming. */
export type PdfReaderHeldState = "saving" | "unsaved";

type WindowSaveStatus = "saved" | "saving" | "unsaved" | "conflict";

const lanes = new Map<string, Map<string, PdfReaderHeldState>>();
const listeners = new Map<string, Set<() => void>>();

function notify(file: string): void {
  for (const listener of [...(listeners.get(file) ?? [])]) listener();
}

/** Say (or, with null, stop saying) that the reader holds `key` unwritten. */
export function holdPdfReaderWork(file: string, key: string, state: PdfReaderHeldState | null): void {
  const lane = lanes.get(file);
  if (state === null) {
    if (!lane?.delete(key)) return;
    if (lane.size === 0) lanes.delete(file);
    notify(file);
    return;
  }
  if (lane?.get(key) === state) return;
  if (lane) lane.set(key, state);
  else lanes.set(file, new Map([[key, state]]));
  notify(file);
}

/** What the reader holds of this file: a failed write outranks a running one. */
export function pdfReaderHeldWork(file: string): PdfReaderHeldState | null {
  const lane = lanes.get(file);
  if (!lane?.size) return null;
  for (const state of lane.values()) if (state === "unsaved") return "unsaved";
  return "saving";
}

/** Whether the reader still holds `key` of this file. */
export function pdfReaderHoldsWork(file: string, key: string): boolean {
  return lanes.get(file)?.has(key) ?? false;
}

export function subscribePdfReaderHeldWork(file: string, listener: () => void): () => void {
  const set = listeners.get(file) ?? new Set<() => void>();
  listeners.set(file, set);
  set.add(listener);
  return () => {
    set.delete(listener);
    if (set.size === 0 && listeners.get(file) === set) listeners.delete(file);
  };
}

const askers = new Map<string, Set<() => void>>();









export function onPdfReaderWorkAsked(file: string, retry: () => void): () => void {
  const set = askers.get(file) ?? new Set<() => void>();
  askers.set(file, set);
  set.add(retry);
  return () => {
    set.delete(retry);
    if (set.size === 0 && askers.get(file) === set) askers.delete(file);
  };
}



export function askPdfReaderWorkNow(file: string): boolean {
  const set = askers.get(file);
  if (!set?.size) return false;
  for (const retry of [...set]) retry();
  return true;
}

/** Tests only. */
export function resetPdfReaderHeldWork(): void {
  lanes.clear();
  listeners.clear();
  askers.clear();
}









export function pdfWindowShownSaveStatus(
  status: WindowSaveStatus,
  held: Readonly<{ reader: PdfReaderHeldState | null; quarantinedRows: number }>,
): WindowSaveStatus {
  if (status !== "saved") return status;
  if (held.quarantinedRows > 0 || held.reader === "unsaved") return "unsaved";
  return held.reader === "saving" ? "saving" : "saved";
}

/** A reader write that failed this many times in a row is said to the user. */
export const PDF_READER_WRITE_SAID_AFTER = 3;

/** The pause before a failed reader write's next try: 1 s, 2 s, 4 s ... 30 s. */
export function pdfReaderRetryPauseMs(failures: number): number {
  const step = Math.max(0, Math.min(5, Math.floor(failures) - 1));
  return Math.min(30_000, 1_000 * 2 ** step);
}
