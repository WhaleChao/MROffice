"use client";














/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { Awareness } from "y-protocols/awareness";

/** Top-level awareness field: `{ d: dataItemId, n: counter }`. */
export const PDF_ANNOTATION_CHANGE_FIELD = "pdfAnn";
/** Writes close together read once on the other desks. */
export const PDF_ANNOTATION_CHANGE_SETTLE_MS = 250;
const MAX_ID_CHARS = 128;

interface ChangeState {
  d: string;
  n: number;
}

function readState(value: unknown): ChangeState | null {
  if (!value || typeof value !== "object") return null;
  const { d, n } = value as { d?: unknown; n?: unknown };
  if (typeof d !== "string" || !d || d.length > MAX_ID_CHARS) return null;
  if (typeof n !== "number" || !Number.isSafeInteger(n) || n < 0) return null;
  return { d, n };
}

/** This desk's write to the data item's annotations landed. */
export function bumpPdfAnnotationChange(awareness: Awareness, dataItemId: string): void {
  if (!dataItemId || dataItemId.length > MAX_ID_CHARS) return;
  const current = readState(awareness.getLocalState()?.[PDF_ANNOTATION_CHANGE_FIELD]);
  const n = current && current.d === dataItemId ? current.n + 1 : 1;
  awareness.setLocalStateField(PDF_ANNOTATION_CHANGE_FIELD, { d: dataItemId, n });
}

/**
 * Calls `onChange` (settled) when another member's counter for the data item
 * moves. Counters already there when watching starts are taken as read (the
 * viewer has just read the lists); a member who appears later with a counter
 * may have written meanwhile, so that reads once too.
 */
export function watchPdfAnnotationChanges(
  awareness: Awareness,
  dataItemId: string,
  onChange: () => void,
): () => void {
  const seen = new Map<number, number>();
  let timer: ReturnType<typeof setTimeout> | null = null;
  const scan = (initial: boolean): boolean => {
    let moved = false;
    awareness.getStates().forEach((state, clientId) => {
      if (clientId === awareness.clientID) return;
      const change = readState(state?.[PDF_ANNOTATION_CHANGE_FIELD]);
      if (!change || change.d !== dataItemId) return;
      if (!initial && seen.get(clientId) !== change.n) moved = true;
      seen.set(clientId, change.n);
    });
    return moved;
  };
  scan(true);
  const handler = () => {
    if (!scan(false) || timer !== null) return;
    timer = setTimeout(() => {
      timer = null;
      onChange();
    }, PDF_ANNOTATION_CHANGE_SETTLE_MS);
  };
  awareness.on("change", handler);
  return () => {
    awareness.off("change", handler);
    if (timer !== null) clearTimeout(timer);
    timer = null;
  };
}
