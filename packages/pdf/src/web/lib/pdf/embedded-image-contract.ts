/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { EmbeddedImageOverlayMutationPlan } from "./embedded-image-apply";
import { PDF_EDITABLE_MAX_BYTES } from "@mrnote/core/web/lib/document-size-limits.ts";

export const MAX_EMBEDDED_IMAGE_OVERLAY_EDITS = 200;

export const MAX_EMBEDDED_IMAGE_REPLACEMENT_BYTES = 16 * 1024 * 1024;

export const MAX_EMBEDDED_IMAGE_REPLACEMENT_PIXELS = 40_000_000;

export const MAX_EMBEDDED_IMAGE_REPLACEMENT_AGGREGATE_BYTES = 64 * 1024 * 1024;

export const MAX_EMBEDDED_IMAGE_OVERLAY_OUTPUT_BYTES = PDF_EDITABLE_MAX_BYTES;

export function embeddedImageOverlayPlanParts(
  plan: EmbeddedImageOverlayMutationPlan,
): readonly string[] {
  if (plan.version !== "mrai.embedded-image-overlay-plan.v1") {
    throw new Error("Invalid embedded-image overlay plan");
  }
  return Object.freeze([
    JSON.stringify({
      version: plan.version,
      sourceByteLength: plan.sourceByteLength,
      pageCount: plan.pageCount,
      targetPages: plan.targetPages,
      replacementCount: plan.replacementCount,
      replacementByteLength: plan.replacementByteLength,
      geometriesDigest: plan.geometriesDigest,
      custodyDigest: plan.custodyDigest,
      formDigest: plan.formDigest,
      metadataDigest: plan.metadataDigest,
      catalogMetadataCustody: plan.catalogMetadataCustody,
    }),
    ...plan.edits.map((edit) => JSON.stringify(edit)),
  ]);
}
