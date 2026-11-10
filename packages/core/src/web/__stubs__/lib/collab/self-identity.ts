/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// lib/collab/self-identity.ts is part of MR Note, not of the open-source editors. This
// stand-in exports what the editors import from it and does nothing; a host
// that has the feature provides it through the host adapter instead.
/* eslint-disable */
import { inertClass, inertComponent, inertFunction, inertHook, inertValue } from "@mrnote/core/host/inert.ts";

export const cachedSelfCollabIdentity: any = inertFunction("lib/collab/self-identity.ts#cachedSelfCollabIdentity", "null");
export type cachedSelfCollabIdentity = any;
export const fetchSelfCollabIdentity: any = inertFunction("lib/collab/self-identity.ts#fetchSelfCollabIdentity", "promise");
export type fetchSelfCollabIdentity = any;
export const useSelfCollabIdentity: any = inertHook("lib/collab/self-identity.ts#useSelfCollabIdentity", "unknown");
export type useSelfCollabIdentity = any;
