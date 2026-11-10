"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  apiPost,
  buildPresignedUploadInit,
  uploadToPresignedUrl,
} from "./api";
import {
  SINGLE_REQUEST_UPLOAD_MAX_BYTES,
  sourceFileFromFile,
  uploadEditorFileInParts,
} from "./document-import/sourceUpload";
import { ensureKnowledgeDataset } from "./knowledge-upload";

export { SINGLE_REQUEST_UPLOAD_MAX_BYTES } from "./document-import/sourceUpload";

type UploadPresignResponse = {
  upload_id: string;
  put_url: string;
  headers: Record<string, string>;
  fields: Record<string, string>;
  upload_method: "PUT" | "POST";
  data_item_id: string;
};

async function uploadOne(datasetId: string, file: File): Promise<string> {
  const presign = await apiPost<UploadPresignResponse>(
    "/api/v1/uploads/presign",
    {
      dataset_id: datasetId,
      filename: file.name,
      media_type: file.type || "application/octet-stream",
      size_bytes: file.size,
    },
  );

  const putRes = await uploadToPresignedUrl(
    presign.put_url,
    buildPresignedUploadInit(presign, file),
    { authenticated: true },
  );
  if (!putRes.ok) {
    throw new Error("upload_failed");
  }

  await apiPost("/api/v1/uploads/complete", {
    upload_id: presign.upload_id,
    data_item_id: presign.data_item_id,
  });
  return presign.data_item_id;
}

export async function uploadNotebookFile(
  notebookId: string,
  file: File,
  options: { readonly signal?: AbortSignal; readonly onProgress?: (sent: number, total: number) => void } = {},
): Promise<string> {
  if (file.size > SINGLE_REQUEST_UPLOAD_MAX_BYTES) {
    const stored = await uploadEditorFileInParts(notebookId, sourceFileFromFile(file), options);
    return stored.data_item_id;
  }
  const dataset = await ensureKnowledgeDataset(notebookId);
  return uploadOne(dataset.id, file);
}
