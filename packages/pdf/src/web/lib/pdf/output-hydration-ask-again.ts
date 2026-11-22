/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useEffect, useRef } from "react";

/**
 * The waits before the marks a PDF's output is checked against (its
 * redactions, its visual covers) are read again by themselves, after a read
 * the server could not answer. The read itself has already been asked a few
 * times over some seconds (annotation-read-batch); these are for a server
 * away longer -- a restart, a database with no connection to give.
 */
export const PDF_OUTPUT_HYDRATION_ASK_AGAIN_WAITS_MS = [5_000, 15_000, 30_000, 60_000, 120_000] as const;

















export function usePdfOutputHydrationAskedAgain(input: Readonly<{
  /** Names the open document; another one starts over. */
  documentIdentity: string;
  /** A read of the marks stands failed. */
  failed: boolean;
  /** Every read of the marks stands answered. */
  ready: boolean;
  /** Changes with every new read of the marks (their generations). */
  round: number;
  /** Whether every failure standing is one the server may answer later. */
  momentary: () => boolean;
  /** Asks the failed reads again (the notice's own retries). */
  askAgain: () => void;
  /** The marks were read after a failure: the other lists are due too. */
  onReadAgain: () => void;
}>): void {
  const { documentIdentity, failed, ready, round } = input;
  const latest = useRef(input);
  // (Before the effect below, every render: it reads the callbacks of the
  // render it runs after.)
  useEffect(() => {
    latest.current = input;
  });
  const state = useRef({ documentIdentity, asked: 0, failedForAMoment: false });
  useEffect(() => {
    if (state.current.documentIdentity !== documentIdentity) {
      state.current = { documentIdentity, asked: 0, failedForAMoment: false };
    }
    const now = state.current;
    if (!failed) {
      if (ready && now.failedForAMoment) {
        now.failedForAMoment = false;
        now.asked = 0;
        latest.current.onReadAgain();
      }
      return;
    }
    if (!latest.current.momentary()) return;
    now.failedForAMoment = true;
    const wait = PDF_OUTPUT_HYDRATION_ASK_AGAIN_WAITS_MS[now.asked];
    if (wait === undefined) return;
    const timer = setTimeout(() => {
      if (state.current !== now) return;
      now.asked += 1;
      latest.current.askAgain();
    }, wait);
    return () => clearTimeout(timer);
  }, [documentIdentity, failed, ready, round]);
}
