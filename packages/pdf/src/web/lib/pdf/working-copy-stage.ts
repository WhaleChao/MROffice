/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * An edit that is only ever written on the working copy (a layer move of
 * several steps, written as one revision and one Undo step: a direct write
 * records one object per revision) when the window has no copy that takes
 * it at the moment it is asked.
 *
 * Why a copy does not take an edit (PdfPageWindow openNativeWorkingCopy /
 * runNativeWorkingEdit), and what is done about it:
 *
 * - `write-in-flight`: another write of the window is still running, or its
 *   journal is busy with one. Waited for, then the copy is opened (made when
 *   there is none) and the edit staged on it.
 * - `staging`: a line being staged is moving the copy. Waited for, likewise.
 * - `runtime-pending`: the window is still reading the saved version it
 *   shows (right after a save, a reload, a version another window saved).
 *   Waited for, likewise: the copy is made on it once it is read.
 * - `not-writable`: the page is read-only here. Refused with that reason.
 * - `unsafe-source`: the source's structure keeps it read-only for writers.
 *   Refused with that reason.
 * - `copy-failed`: the copy failed (a save of it refused for good) or was
 *   closed. Refused: saving or reopening the page starts a new one.
 * - `policy`: the command's own policy refuses it for good. Refused with
 *   the policy's reason.
 */

export type PdfWorkingCopyMiss =
  | "write-in-flight"
  | "staging"
  | "runtime-pending"
  | "not-writable"
  | "unsafe-source"
  | "copy-failed"
  | "policy";

/** Whether waiting can make the copy take the edit. */
export function pdfWorkingCopyMissIsTransient(miss: PdfWorkingCopyMiss | null): boolean {
  return miss === "write-in-flight" || miss === "staging" || miss === "runtime-pending";
}

/** The words for an edit no copy took (pdf.pageWindow.workingCopyMiss.*). */
export function pdfWorkingCopyMissMessageKey(miss: PdfWorkingCopyMiss | null): string {
  switch (miss) {
    case "not-writable":
      return "pdf.pageWindow.workingCopyMiss.notWritable";
    case "unsafe-source":
      return "pdf.pageWindow.workingCopyMiss.unsafeSource";
    case "copy-failed":
      return "pdf.pageWindow.workingCopyMiss.copyFailed";
    case "policy":
      return "pdf.pageWindow.workingCopyMiss.policy";
    case "write-in-flight":
    case "staging":
    case "runtime-pending":
      return "pdf.pageWindow.workingCopyMiss.busy";
    default:
      return "pdf.pageWindow.workingCopyMiss.unknown";
  }
}

export type PdfWorkingCopyStageResult<T> =
  | Readonly<{ staged: true; value: T }>
  | Readonly<{ staged: false; miss: PdfWorkingCopyMiss | null }>;

/**
 * Stage an edit on the working copy, waiting out what keeps the copy from
 * taking it for a while (see the module comment) and trying again; refused,
 * with why, once what is left is no reason to wait, or `deadlineMs` passed.
 * `attempt` resolves null when no copy took the edit (`lastMiss` then says
 * why), else the staged result; `settle` waits for what the miss names.
 */
export async function stageOnPdfWorkingCopy<T>(input: {
  readonly attempt: () => Promise<T | null>;
  readonly lastMiss: () => PdfWorkingCopyMiss | null;
  readonly settle: (miss: PdfWorkingCopyMiss) => Promise<void>;
  readonly deadlineMs?: number;
  readonly now?: () => number;
}): Promise<PdfWorkingCopyStageResult<T>> {
  const now = input.now ?? (() => Date.now());
  const deadline = now() + (input.deadlineMs ?? 20_000);
  for (;;) {
    const value = await input.attempt();
    if (value !== null) return Object.freeze({ staged: true as const, value });
    const miss = input.lastMiss();
    if (!pdfWorkingCopyMissIsTransient(miss) || now() >= deadline) {
      return Object.freeze({ staged: false as const, miss });
    }
    await input.settle(miss!);
  }
}
