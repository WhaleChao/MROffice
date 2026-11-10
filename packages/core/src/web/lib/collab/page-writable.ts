/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// Who may write a page is the service's to say; a local file is the user's
// own, so it is always writable here.
import { createElement, type ReactNode } from "react";

export * from "@mrnote/core/web/__stubs__/lib/collab/page-writable.ts";

export interface PageWritable {
  writable: boolean;
  reason: string | null;
  lockedByName: string | null;
  known: boolean;
}

const WRITABLE: PageWritable = Object.freeze({ writable: true, reason: null, lockedByName: null, known: true });

export const PAGE_AUTHORITY_OUTAGE_GRACE_MS = 60_000;
export const PAGE_AUTHORITY_CHANGED_EVENT = "mrn:page-authority-changed";

export function usePageWritable(..._args: unknown[]): PageWritable {
  return WRITABLE;
}
export function usePageWritableThroughOutages(..._args: unknown[]): PageWritable {
  return WRITABLE;
}
export function invalidatePageWritable(..._args: unknown[]): void {}
export function usePageOwner(..._args: unknown[]): null {
  return null;
}
export function cachedPageVisibility(..._args: unknown[]): null {
  return null;
}
export function cachedPageOwner(..._args: unknown[]): null {
  return null;
}
export function usePageWorkspaceId(): string | null {
  return null;
}
export function PageWorkspaceScopeProvider({ children }: { children?: ReactNode }) {
  return createElement("div", { style: { display: "contents" } }, children);
}
