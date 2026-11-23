"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useEffect, useRef } from "react";

import { reportOfficeExperience } from "./office-experience-report";

export const ROOM_DURABLE_STALL_MS = 60_000;
export const SAVE_SURFACE_STALL_MS = 90_000;

const reportedThisSession = new Set<string>();

/** Test seam. */
export function resetSaveStallReportsForTests(): void {
  reportedThisSession.clear();
}

/** One report per kind and page in this session. Returns whether it went out. */
export function reportSaveStallOnce(
  kind: "room_durable_stalled" | "save_surface_stalled",
  pageId: string,
  pageType: string | null | undefined,
  code: string,
): boolean {
  const key = `${kind}|${pageId}`;
  if (reportedThisSession.has(key)) return false;
  const sent = reportOfficeExperience(kind, { pageId, pageType, code });
  if (sent) reportedThisSession.add(key);
  return sent;
}

/** The inputs of a face as one code (`face:name.value:...`, booleans 0/1), in the report code alphabet. */
export function saveStallCode(face: string, inputs: Readonly<Record<string, string | number | boolean | null | undefined>>): string {
  const parts = [face.replace(/[^A-Za-z0-9_.-]/g, "") || "face"];
  for (const [key, value] of Object.entries(inputs)) {
    const shown = value === true ? "1" : value === false || value == null ? "0" : String(value).replace(/[^A-Za-z0-9_.-]/g, "");
    parts.push(`${key}.${shown}`);
  }
  return parts.join(":").slice(0, 96);
}

/**
 * Reports a face that has said `saving` or `unsaved` for
 * `SAVE_SURFACE_STALL_MS` with no request on its way (`busy` false), once per
 * page in the session. `inputs` is read when it fires.
 */
export function useSaveSurfaceStallReport(options: {
  pageId: string | null | undefined;
  pageType: string | null | undefined;
  face: string;
  busy: boolean;
  enabled: boolean;
  inputs: () => Readonly<Record<string, string | number | boolean | null | undefined>>;
}): void {
  const inputsRef = useRef(options.inputs);
  inputsRef.current = options.inputs;
  const { pageId, pageType, face, busy, enabled } = options;
  const waiting = enabled && !busy && Boolean(pageId) && (face === "saving" || face === "unsaved");
  useEffect(() => {
    if (!waiting || !pageId) return undefined;
    const timer = window.setTimeout(() => {
      try {
        reportSaveStallOnce("save_surface_stalled", pageId, pageType, saveStallCode(face, inputsRef.current()));
      } catch {
        // A report never gets in the way of saving.
      }
    }, SAVE_SURFACE_STALL_MS);
    return () => window.clearTimeout(timer);
  }, [face, pageId, pageType, waiting]);
}

/** A room's durable confirmations, watched: `PageRoomDurableStalled` once the
 *  state has stayed active and unsettled for `ROOM_DURABLE_STALL_MS`. Returns
 *  the stop function. */
export function watchRoomDurableStall(options: {
  pageId: string;
  pageType: string | null | undefined;
  read: () => { active: boolean; settled: boolean; refusal: string | null };
  subscribe: (listener: () => void) => () => void;
}): () => void {
  let timer: number | null = null;
  const check = () => {
    let state: { active: boolean; settled: boolean; refusal: string | null };
    try { state = options.read(); } catch { return; }
    const stalled = state.active && !state.settled && state.refusal === null;
    if (!stalled) {
      if (timer !== null) window.clearTimeout(timer);
      timer = null;
      return;
    }
    if (timer !== null) return;
    timer = window.setTimeout(() => {
      timer = null;
      let now: { active: boolean; settled: boolean; refusal: string | null };
      try { now = options.read(); } catch { return; }
      if (now.active && !now.settled && now.refusal === null) {
        reportSaveStallOnce("room_durable_stalled", options.pageId, options.pageType, "unconfirmed_60s");
      }
    }, ROOM_DURABLE_STALL_MS);
  };
  check();
  const stop = options.subscribe(check);
  return () => {
    stop();
    if (timer !== null) window.clearTimeout(timer);
  };
}

type StallInputs = Record<string, string | number | boolean | null | undefined>;

/** What a face that is not waiting says, for the stall watch: it never fires. */
const SETTLED = "settled";

/** Docs (NoteEditor's save face): the face it shows and the inputs that make it. */
export function docsSaveStallFace(input: {
  saveStatus: string; partsViewStopped: boolean; hasConflict: boolean; syncedToServer: boolean;
  notConnected: boolean; error: boolean; raw: string; catchingUp: boolean; settling: boolean;
  closingEdits: boolean; partsSaving: boolean; local: string; conflictCopy: boolean; dirty: boolean;
}): { face: string; inputs: StallInputs } {
  const face = input.partsViewStopped || input.hasConflict || input.notConnected
    || (input.syncedToServer && input.saveStatus === "unsaved") || (input.error && input.saveStatus === "unsaved")
    ? SETTLED : input.saveStatus;
  return { face, inputs: { raw: input.raw, cu: input.catchingUp, st: input.settling, ce: input.closingEdits,
    ps: input.partsSaving, local: input.local, cc: input.conflictCopy, dirty: input.dirty } };
}

/** Sheets (SpreadsheetWindow's save face). */
export function sheetSaveStallFace(input: {
  saveStatus: string; hasConflict: boolean; syncedToServer: boolean; notConnected: boolean; error: boolean;
  saveSlow: boolean; dirty: boolean; catchingUp: boolean; roomsAhead: boolean; retrying: string | false;
}): { face: string; inputs: StallInputs } {
  const face = input.hasConflict || input.notConnected || (input.syncedToServer && input.saveStatus === "unsaved")
    || (input.error && input.saveStatus === "unsaved") ? SETTLED : input.saveSlow ? "saving" : input.saveStatus;
  return { face, inputs: { raw: input.saveStatus, slow: input.saveSlow, cu: input.catchingUp, ra: input.roomsAhead,
    gr: input.retrying, dirty: input.dirty } };
}

/** PDF (PdfPageWindow's save face). */
export function pdfSaveStallFace(input: {
  shown: string; raw: string; textSaveFailed: boolean; syncedToServer: boolean; reader: string | null;
  textBoxes: string | null; carriedWords: string | null; copyWrites: number; quarantinedRows: number; dirty: boolean;
}): { face: string; inputs: StallInputs } {
  const face = input.textSaveFailed || input.syncedToServer ? SETTLED : input.shown;
  return { face, inputs: { raw: input.raw, rd: input.reader ?? "none", tb: input.textBoxes ?? "none",
    cw: input.carriedWords ?? "none", cp: input.copyWrites, q: input.quarantinedRows, dirty: input.dirty } };
}
