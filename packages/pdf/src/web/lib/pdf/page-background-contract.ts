/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PageBackgroundMutationPlan } from "./page-background";
import { PDF_EDITABLE_MAX_BYTES } from "@mrnote/core/web/lib/document-size-limits.ts";

export const MAX_PAGE_BACKGROUND_IMAGE_BYTES = 16 * 1024 * 1024;

export const MAX_PAGE_BACKGROUND_IMAGE_PIXELS = 40_000_000;

export const MAX_PAGE_BACKGROUND_TARGETS = 10_000;

export const MAX_PAGE_BACKGROUND_OUTPUT_BYTES = PDF_EDITABLE_MAX_BYTES;

export function pageBackgroundPlanParts(
  plan: PageBackgroundMutationPlan,
): readonly string[] {
  if (plan.version !== "mrai.page-background-plan.v1") {
    throw new Error("Invalid PDF background plan");
  }
  return Object.freeze([JSON.stringify(plan)]);
}
