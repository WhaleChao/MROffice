/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// lib/collab/ws-provider.ts is part of MR Note, not of the open-source editors. This
// stand-in exports what the editors import from it and does nothing; a host
// that has the feature provides it through the host adapter instead.
/* eslint-disable */
import { inertClass, inertComponent, inertFunction, inertHook, inertValue } from "@mrnote/core/host/inert.ts";

export const COLLAB_INITIAL_SYNC_BUDGET_MS: any = 12_000;
export type COLLAB_INITIAL_SYNC_BUDGET_MS = any;
export type CollabNetworkState = any;
export type PageYjsFailureCause = any;
export const collabSocketOrigin: any = inertFunction("lib/collab/ws-provider.ts#collabSocketOrigin", "null");
export type collabSocketOrigin = any;
export const createPageCausalModelStatus: any = inertFunction("lib/collab/ws-provider.ts#createPageCausalModelStatus", "unknown");
export type createPageCausalModelStatus = any;
export const getPageYjsTransportSnapshot: any = inertFunction("lib/collab/ws-provider.ts#getPageYjsTransportSnapshot", "unknown");
export type getPageYjsTransportSnapshot = any;
export const prepareWhiteboardObjectEditing: any = inertFunction("lib/collab/ws-provider.ts#prepareWhiteboardObjectEditing", "promise-void");
export type prepareWhiteboardObjectEditing = any;
export const ticketedWebsocketProviderFactory: any = inertFunction("lib/collab/ws-provider.ts#ticketedWebsocketProviderFactory", "unknown");
export type ticketedWebsocketProviderFactory = any;
