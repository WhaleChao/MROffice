"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useCallback, useSyncExternalStore } from "react";
import { pdfMarkupMarksPending, subscribePdfMarkupWrites } from "./markup-write-queue";




export function usePdfMarkupMarksPending(key: string | null | undefined): boolean {
  const subscribe = useCallback(
    (listener: () => void) => (key ? subscribePdfMarkupWrites(key, listener) : () => undefined),
    [key],
  );
  const read = useCallback(() => (key ? pdfMarkupMarksPending(key) : false), [key]);
  return useSyncExternalStore(subscribe, read, () => false);
}
