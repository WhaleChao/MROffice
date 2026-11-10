"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Upload a file as an attachment of a page, whatever its size.
 *
 * A file up to {@link SINGLE_REQUEST_UPLOAD_MAX_BYTES} is one multipart
 * request (`POST /pages/{id}/attachments/upload`); a larger one goes up in
 * parts and is attached when they have all arrived
 * (`POST /source-uploads/{id}/attach`), up to the server's 5 GiB, so the
 * request-body bound no longer decides how large an attachment may be. Both
 * answer the same attachment fields.
 */

import { apiPostFormData } from "./api";
import {
  sourceFileFromFile,
  uploadPageAttachmentInParts,
  type UploadedPageAttachment,
} from "./document-import/sourceUpload";
import { notebookSDK } from "./notebook-sdk";
import { SINGLE_REQUEST_UPLOAD_MAX_BYTES } from "./study-upload";

export type { UploadedPageAttachment };

export interface PageAttachmentUploadOptions {
  readonly signal?: AbortSignal;
  readonly title?: string;
  /** The page's notebook, when the caller knows it (saves one lookup). */
  readonly notebookId?: string | null;
  readonly onProgress?: (sentBytes: number, totalBytes: number) => void;
}

async function notebookOfPage(pageId: string): Promise<string> {
  const page = await notebookSDK.getPage(pageId);
  if (!page.notebook_id) throw new Error("Notebook is not ready for attachments yet.");
  return page.notebook_id;
}

export async function uploadPageAttachment(
  pageId: string,
  file: File,
  options: PageAttachmentUploadOptions = {},
): Promise<UploadedPageAttachment> {
  if (file.size > SINGLE_REQUEST_UPLOAD_MAX_BYTES) {
    const notebookId = options.notebookId || (await notebookOfPage(pageId));
    return uploadPageAttachmentInParts(notebookId, pageId, sourceFileFromFile(file), {
      signal: options.signal,
      title: options.title,
      onProgress: options.onProgress,
    });
  }
  const formData = new FormData();
  formData.append("file", file);
  if (options.title) formData.append("title", options.title);
  return apiPostFormData<UploadedPageAttachment>(
    `/api/v1/pages/${encodeURIComponent(pageId)}/attachments/upload`,
    formData,
    { signal: options.signal },
  );
}
