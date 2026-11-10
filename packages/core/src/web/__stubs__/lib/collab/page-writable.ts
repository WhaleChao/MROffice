/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// lib/collab/page-writable.ts is part of MR Note, not of the open-source editors. This
// stand-in exports what the editors import from it and does nothing; a host
// that has the feature provides it through the host adapter instead.
/* eslint-disable */
import { inertClass, inertComponent, inertFunction, inertHook, inertValue } from "@mrnote/core/host/inert.ts";

export const PAGE_AUTHORITY_OUTAGE_GRACE_MS: any = 60_000;
export type PAGE_AUTHORITY_OUTAGE_GRACE_MS = any;
export type PageWritable = any;
export const publishPageCollabAuthority: any = inertFunction("lib/collab/page-writable.ts#publishPageCollabAuthority", "void");
export type publishPageCollabAuthority = any;
export const useLastVerifiedPageCollabBinding: any = inertHook("lib/collab/page-writable.ts#useLastVerifiedPageCollabBinding", "unknown");
export type useLastVerifiedPageCollabBinding = any;
export const usePageWorkspaceId: any = inertHook("lib/collab/page-writable.ts#usePageWorkspaceId", "null");
export type usePageWorkspaceId = any;
export const usePageWritable: any = inertHook("lib/collab/page-writable.ts#usePageWritable", "unknown");
export type usePageWritable = any;
export const usePageWritableThroughOutages: any = inertHook("lib/collab/page-writable.ts#usePageWritableThroughOutages", "unknown");
export type usePageWritableThroughOutages = any;
