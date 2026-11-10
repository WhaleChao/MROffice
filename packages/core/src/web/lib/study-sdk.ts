/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Resource projection retained for the PDF/reference-document window. */

import { apiGet } from "./api";

export type StudyAssetStatus = "pending" | "parsing" | "chunked" | "indexed" | "failed";

/** Mirrors `StudyAssetOut` in apps/api/app/schemas/study.py. */
export interface StudyAsset {
  id: string;
  notebook_id: string;
  data_item_id: string | null;
  title: string;
  asset_type: string;
  status: StudyAssetStatus;
  total_chunks: number;
  metadata_json: Record<string, unknown> | null;
  tags: string[];
  created_at: string;
  updated_at: string;
}

export function listAssets(
  notebookId: string,
): Promise<{ items: StudyAsset[]; total: number }> {
  return apiGet<{ items: StudyAsset[]; total: number }>(
    `/api/v1/notebooks/${notebookId}/study-assets`,
  );
}

// ---------------------------------------------------------------------------
// Grouped export
// ---------------------------------------------------------------------------

export const studySDK = {
  listAssets,
} as const;

export default studySDK;
