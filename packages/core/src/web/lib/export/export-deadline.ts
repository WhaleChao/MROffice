"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { runCollaborationRequestWithTimeout } from "../collaboration-windows/request-timeout";

/**
 * Bounds on the waits one editor export makes before the file reaches the
 * browser. An export holds the window's "export running" latch (its menu
 * items are disabled and a second export is told one is running) for as
 * long as it waits, so a wait that never settles left export unusable until
 * the page was reloaded, with nothing said (xlsx-rounds R490-X01 session A:
 * Export > XLSX on a tiled workbook just past a row insert sent no request
 * for 180 s). Each wait is bounded; past its bound the export fails closed
 * with words the person can read and a way to try again, and the latch is
 * released.
 */

/** Making the version the export binds: write authority, the draft's save. */
export const EXPORT_PREPARE_DEADLINE_MS = 120_000;
/**
 * The export request and its file: the server answers an export within its
 * own 300 s request bound, so a wait past this one is one that will not end.
 */
export const EXPORT_TRANSFER_DEADLINE_MS = 360_000;
/** Times an export waits for the page's authority and runs again, at most. */
export const EXPORT_AUTHORITY_RERUNS = 3;

/** A bounded export wait that ran out: said in `message`, safe to retry. */
export class ExportDeadlineError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "ExportDeadlineError";
  }
}

/** The window went away while its export waited: nothing is said. */
export class ExportCancelledError extends Error {
  constructor() {
    super("The export was cancelled.");
    this.name = "ExportCancelledError";
  }
}

export function isExportDeadlineError(error: unknown): error is ExportDeadlineError {
  return error instanceof ExportDeadlineError;
}

export function isExportCancelledError(error: unknown): error is ExportCancelledError {
  return error instanceof ExportCancelledError;
}

/**
 * Run one export wait under `controller` (one controller per export
 * attempt). Past `timeoutMs` the controller is aborted -- the request it
 * carries is cancelled and every later wait of the attempt fails at once --
 * and this rejects with an {@link ExportDeadlineError} saying `message`. An
 * attempt cancelled with an {@link ExportCancelledError} rejects with it.
 */
export async function withinExportDeadline<T>(
  controller: AbortController,
  work: (signal: AbortSignal) => Promise<T>,
  message: string,
  timeoutMs: number,
): Promise<T> {
  try {
    return await runCollaborationRequestWithTimeout(controller, work, message, timeoutMs);
  } catch (error) {
    const reason: unknown = controller.signal.aborted ? controller.signal.reason : undefined;
    if (isExportCancelledError(reason)) throw reason;
    if (reason !== undefined && error === reason) throw new ExportDeadlineError(message);
    throw error;
  }
}
