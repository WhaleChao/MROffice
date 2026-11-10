/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// Who may write in a workspace is the service's to say; a local file is the
// user's own, so the editors' own role checks answer as they do for an owner
// and the workspace hook (which asks the service) reports a workspace of one.

export * from "@mrnote/core/web/__stubs__/lib/use-workspace-me.ts";

export const WORKSPACE_PRIVILEGED_ROLES: ReadonlySet<string> = new Set(["owner", "admin"]);

export function isWorkspacePrivilegedRole(role: string | undefined): boolean {
  return Boolean(role && WORKSPACE_PRIVILEGED_ROLES.has(role));
}

export const WORKSPACE_WRITE_ROLES = Object.freeze(["owner", "admin", "editor"] as const);
const WORKSPACE_WRITE_ROLE_SET: ReadonlySet<string> = new Set(WORKSPACE_WRITE_ROLES);

export function isWorkspaceWriteRole(role: string | null | undefined): boolean {
  return typeof role === "string" && WORKSPACE_WRITE_ROLE_SET.has(role);
}

const LOCAL_WORKSPACE = Object.freeze({
  data: Object.freeze({ workspace_id: "local", enable_universal_editor: false, role: "owner", feature_flags: Object.freeze({}) }),
  loading: false,
  error: null,
  refresh: () => {},
});

export function useWorkspaceMe(..._args: unknown[]): typeof LOCAL_WORKSPACE {
  return LOCAL_WORKSPACE;
}

export function resetWorkspaceMeCache(): void {}
