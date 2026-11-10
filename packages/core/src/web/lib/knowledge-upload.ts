"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { apiGet, apiPost, buildPresignedUploadInit, uploadToPresignedUrl } from "./api";
import {
  SINGLE_REQUEST_UPLOAD_MAX_BYTES,
  sourceFileFromFile,
  uploadEditorFileInParts,
} from "./document-import/sourceUpload";

type DatasetInfo = {
  id: string;
  name: string;
  type: string;
};

type UploadPresignResponse = {
  upload_id: string;
  put_url: string;
  headers: Record<string, string>;
  fields: Record<string, string>;
  upload_method: "PUT" | "POST";
  data_item_id: string;
};

export async function ensureKnowledgeDataset(notebookId: string): Promise<DatasetInfo> {
  const datasets = await apiGet<DatasetInfo[]>(
    `/api/v1/datasets?notebook_id=${encodeURIComponent(notebookId)}`,
  );
  if (datasets.length > 0) {
    return datasets[0];
  }

  return apiPost<DatasetInfo>("/api/v1/datasets", {
    notebook_id: notebookId,
    name: "Default Knowledge",
    type: "text",
  });
}

async function uploadSingleFile(notebookId: string, datasetId: string, file: File): Promise<void> {
  if (file.size > SINGLE_REQUEST_UPLOAD_MAX_BYTES) {
    await uploadEditorFileInParts(notebookId, sourceFileFromFile(file));
    return;
  }
  const presign = await apiPost<UploadPresignResponse>("/api/v1/uploads/presign", {
    dataset_id: datasetId,
    filename: file.name,
    media_type: file.type || "application/octet-stream",
    size_bytes: file.size,
  });

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
}

export async function uploadKnowledgeFiles(notebookId: string, files: File[]): Promise<void> {
  if (files.length === 0) {
    return;
  }

  const dataset = await ensureKnowledgeDataset(notebookId);
  for (const file of files) {
    await uploadSingleFile(notebookId, dataset.id, file);
  }
}
