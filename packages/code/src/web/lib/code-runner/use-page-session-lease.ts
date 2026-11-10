"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useEffect, useRef } from "react";

import { dispatchNotebookProposalsConfirmed } from "@mrnote/core/web/lib/notebook-events.ts";
import {
  acquirePageRuntime,
  type RuntimeLease,
} from "./page-runtime-lease";

interface PageVersionObservation {
  pageId: string;
  version: number;
  reloadNonce: number;
}

export interface UsePageSessionLeaseOptions {
  pageId: string;
  /** Tear down page-owned workers/databases after the final view closes. */
  dispose: () => void;
  /** Latest server content version exposed by the autosave hook. */
  contentVersion?: number | null;
  /** Increments on hydration/reload, but not on this hook's own save. */
  reloadNonce: number;
}

/**
 * Own a page-scoped runtime lease and keep duplicate views on one safe save
 * version.
 *
 * A normal autosave advances `contentVersion` without advancing `reloadNonce`.
 * When that happens while another view holds the same page lease, publish the
 * existing targeted external-version event. Every autosave hook already handles
 * that event fail-closed: a clean peer rehydrates, while a peer that typed during
 * the fetch keeps its full-body draft and enters the normal CAS conflict path.
 *
 * Reusing that path matters. Calling `reload()` directly would clear a draft
 * typed while its GET was in flight; rebasing the peer blindly would let the
 * stale body overwrite the just-saved one. The proposal-named event is the
 * repository's established signal for the underlying condition: this page's
 * server version advanced outside the receiving hook.
 */
export function usePageSessionLease({
  pageId,
  dispose,
  contentVersion,
  reloadNonce,
}: UsePageSessionLeaseOptions): void {
  const leaseRef = useRef<RuntimeLease | null>(null);
  const previousRef = useRef<PageVersionObservation | null>(null);

  useEffect(() => {
    const lease = acquirePageRuntime(pageId, dispose);
    leaseRef.current = lease;
    return () => {
      if (leaseRef.current === lease) leaseRef.current = null;
      lease.release();
    };
  }, [dispose, pageId]);

  useEffect(() => {
    const lease = leaseRef.current;
    if (!lease || lease.key !== pageId) return;
    if (
      typeof contentVersion !== "number" ||
      !Number.isSafeInteger(contentVersion) ||
      contentVersion < 1
    ) {
      if (previousRef.current?.pageId !== pageId) {
        previousRef.current = null;
      }
      return;
    }

    const previous = previousRef.current;
    if (!previous || previous.pageId !== pageId) {
      previousRef.current = { pageId, version: contentVersion, reloadNonce };
      lease.observeContentVersion(contentVersion);
      return;
    }

    const advanced = contentVersion > previous.version;
    const savedByThisView = advanced && reloadNonce === previous.reloadNonce;
    previousRef.current = {
      pageId,
      version: Math.max(previous.version, contentVersion),
      reloadNonce,
    };

    if (!advanced) {
      lease.observeContentVersion(contentVersion);
      return;
    }
    if (!savedByThisView) {
      // A reload/restore/external write is already the reconciliation itself.
      // Record its version without starting another notification cycle.
      lease.observeContentVersion(contentVersion);
      return;
    }
    if (lease.publishSavedContentVersion(contentVersion)) {
      dispatchNotebookProposalsConfirmed({ documentId: pageId });
    }
  }, [contentVersion, pageId, reloadNonce]);
}
