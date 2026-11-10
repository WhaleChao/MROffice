/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// Who is editing: in MR Note the signed-in account; here, with no accounts,
// the person at this computer, under one local id. Comments and notes they
// write are signed with this name (the editors' own words "Author").
import { useMemo } from "react";

export * from "@mrnote/core/web/__stubs__/lib/collab/self-identity.ts";

export interface SelfCollabIdentity {
  id: string | null;
  name: string | null;
  color: string | null;
}

const LOCAL: SelfCollabIdentity = Object.freeze({ id: "local-author", name: "Author", color: "#2f6fd6" });

export function cachedSelfCollabIdentity(): SelfCollabIdentity | null {
  return LOCAL;
}

export function fetchSelfCollabIdentity(): Promise<SelfCollabIdentity> {
  return Promise.resolve(LOCAL);
}

export function invalidateSelfCollabIdentity(): void {}

export const resetSelfCollabIdentityCache = invalidateSelfCollabIdentity;

export function useSelfCollabIdentity(): { id: string; name: string } {
  return useMemo(() => ({ id: LOCAL.id ?? "", name: LOCAL.name ?? "" }), []);
}
