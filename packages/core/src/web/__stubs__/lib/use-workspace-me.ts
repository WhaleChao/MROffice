/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// lib/use-workspace-me.ts is part of MR Note, not of the open-source editors. This
// stand-in exports what the editors import from it and does nothing; a host
// that has the feature provides it through the host adapter instead.
/* eslint-disable */
import { inertClass, inertComponent, inertFunction, inertHook, inertValue } from "@mrnote/core/host/inert.ts";

export type UseWorkspaceMeResult = any;
export type WorkspaceMe = any;
export const isWorkspaceWriteRole: any = inertFunction("lib/use-workspace-me.ts#isWorkspaceWriteRole", "boolean");
export type isWorkspaceWriteRole = any;
export const useWorkspaceMe: any = inertHook("lib/use-workspace-me.ts#useWorkspaceMe", "unknown");
export type useWorkspaceMe = any;
