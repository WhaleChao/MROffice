/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { apiGet, apiPostFormData } from "@mrnote/core/web/lib/api.ts";

export type SpreadsheetDataFromPictureSource = "file" | "clipboard";

export interface SpreadsheetDataFromPictureCapability {
  readonly available: boolean;
  readonly reasonCode: string | null;
  readonly policyVersion: number | null;
  readonly policyMarkerSha256: string | null;
  readonly modelId: string | null;
  readonly residencyRegion: string | null;
  readonly maxImageBytes: number;
  readonly acceptedMediaTypes: readonly string[];
  readonly sources: readonly SpreadsheetDataFromPictureSource[];
}

export interface SpreadsheetDataFromPictureCell {
  readonly text: string;
  readonly confidence: number;
  readonly reviewRequired: boolean;
}

export interface SpreadsheetDataFromPictureResult {
  readonly source: SpreadsheetDataFromPictureSource;
  readonly rows: readonly (readonly SpreadsheetDataFromPictureCell[])[];
  readonly warnings: readonly string[];
  readonly lowConfidenceThreshold: number;
  readonly policyVersion: number;
  readonly policyMarkerSha256: string;
  readonly modelId: string;
  readonly residencyRegion: string;
}

export async function loadSpreadsheetDataFromPictureCapability(
  pageId: string,
  signal?: AbortSignal,
): Promise<SpreadsheetDataFromPictureCapability> {
  return apiGet<SpreadsheetDataFromPictureCapability>(
    `/api/v1/pages/${encodeURIComponent(pageId)}/spreadsheet/data-from-picture/capability`,
    { signal },
  );
}

export async function analyzeSpreadsheetPicture(
  pageId: string,
  input: {
    readonly file: File;
    readonly source: SpreadsheetDataFromPictureSource;
    readonly privacyAcknowledged: boolean;
    readonly capability: SpreadsheetDataFromPictureCapability;
  },
  signal?: AbortSignal,
): Promise<SpreadsheetDataFromPictureResult> {
  if (
    input.capability.policyVersion === null ||
    input.capability.policyMarkerSha256 === null
  ) {
    throw new Error("Data from Picture authority is unavailable");
  }
  const form = new FormData();
  form.set("file", input.file);
  form.set("source", input.source);
  form.set(
    "privacy_acknowledged",
    input.privacyAcknowledged ? "true" : "false",
  );
  form.set(
    "expected_policy_version",
    String(input.capability.policyVersion),
  );
  form.set(
    "expected_policy_marker_sha256",
    input.capability.policyMarkerSha256,
  );
  return apiPostFormData<SpreadsheetDataFromPictureResult>(
    `/api/v1/pages/${encodeURIComponent(pageId)}/spreadsheet/data-from-picture/analyze`,
    form,
    { signal },
  );
}

export async function readClipboardPicture(): Promise<File> {
  if (
    typeof navigator === "undefined" ||
    typeof navigator.clipboard?.read !== "function"
  ) {
    throw new Error("clipboard_image_unavailable");
  }
  const items = await navigator.clipboard.read();
  for (const item of items) {
    const imageType = item.types.find((type) => type.startsWith("image/"));
    if (!imageType) continue;
    const blob = await item.getType(imageType);
    return new File([blob], "clipboard-image", {
      type: blob.type || imageType,
    });
  }
  throw new Error("clipboard_image_missing");
}
