/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import * as React from "react";


















const OWNER_STACK_CUTOFF = 10_000;

type ReactSharedInternals = { recentlyCreatedOwnerStacks?: unknown };

function sharedInternals(): ReactSharedInternals | null {
  const internals = (
    React as unknown as {
      __CLIENT_INTERNALS_DO_NOT_USE_OR_WARN_USERS_THEY_CANNOT_UPGRADE?: unknown;
    }
  ).__CLIENT_INTERNALS_DO_NOT_USE_OR_WARN_USERS_THEY_CANNOT_UPGRADE;
  return internals && typeof internals === "object"
    ? (internals as ReactSharedInternals)
    : null;
}

export function skipDevOwnerStacksForThisRender(): void {
  if (process.env.NODE_ENV === "production") return;
  const internals = sharedInternals();
  if (
    !internals ||
    typeof internals.recentlyCreatedOwnerStacks !== "number" ||
    internals.recentlyCreatedOwnerStacks >= OWNER_STACK_CUTOFF
  ) {
    return;
  }
  internals.recentlyCreatedOwnerStacks = OWNER_STACK_CUTOFF;
}
