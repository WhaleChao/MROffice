/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The transfer line's store (lib/transfer-store) and its React hook.
 */

import { useEffect, useSyncExternalStore } from "react";

import {
  NO_TRANSFERS,
  beginTransfer,
  currentTransfers,
  subscribeTransfers,
  type TransferSnapshot,
} from "./transfer-store";

export * from "./transfer-store";

// The server value is the store's own empty array, so the first client
// render matches it by identity (a separate [] differed on every load).
const serverTransfers = () => NO_TRANSFERS;

export function useTransfers(): readonly TransferSnapshot[] {
  return useSyncExternalStore(subscribeTransfers, currentTransfers, serverTransfers);
}

/**
 * While ``name`` is set, the transfer line shows it as one operation with
 * its time so far (an export from click to file): the line at the top of
 * the window, not only a chip in the editor's status bar.
 */
export function useTransferWhile(name: string | null | undefined): void {
  const active = Boolean(name);
  useEffect(() => {
    if (!active || !name) return;
    const started = Date.now();
    const transfer = beginTransfer({ direction: "download", name, totalBytes: 0, umbrella: "primary" });
    const timer = window.setInterval(() => transfer.preparing((Date.now() - started) / 1000), 1000);
    return () => {
      window.clearInterval(timer);
      transfer.end();
    };
    // The name is read once: a re-render with new wording is the same export.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [active]);
}
