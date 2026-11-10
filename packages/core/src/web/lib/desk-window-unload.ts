"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useEffect, useRef } from "react";


















const probes = new Map<string, () => boolean>();

export function deskWindowCanUnload(windowId: string): boolean {
  const probe = probes.get(windowId);
  if (!probe) return false;
  try {
    return probe();
  } catch {
    return false;
  }
}

/** Register the editor of desk window `windowId` (none for an editor that is
 * not a desk window's: embedded, a guest page). `canUnload` is read when the
 * desk asks, always the latest one. */
export function useDeskWindowUnloadable(
  windowId: string | undefined,
  canUnload: () => boolean,
): void {
  const latest = useRef(canUnload);
  useEffect(() => {
    latest.current = canUnload;
  });
  useEffect(() => {
    if (!windowId) return;
    const probe = () => latest.current();
    probes.set(windowId, probe);
    return () => {
      if (probes.get(windowId) === probe) probes.delete(windowId);
    };
  }, [windowId]);
}
