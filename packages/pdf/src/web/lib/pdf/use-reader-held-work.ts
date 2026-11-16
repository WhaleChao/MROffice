"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useCallback, useSyncExternalStore } from "react";
import {
  pdfReaderHeldWork,
  subscribePdfReaderHeldWork,
  type PdfReaderHeldState,
} from "./reader-held-work";

/** What the reader holds of this file that the file does not hold yet (the
 * window reads 保存中 or 未保存 over it, and asks before closing). */
export function usePdfReaderHeldWork(file: string | null | undefined): PdfReaderHeldState | null {
  const subscribe = useCallback(
    (listener: () => void) => (file ? subscribePdfReaderHeldWork(file, listener) : () => undefined),
    [file],
  );
  const read = useCallback(() => (file ? pdfReaderHeldWork(file) : null), [file]);
  return useSyncExternalStore(subscribe, read, () => null);
}
