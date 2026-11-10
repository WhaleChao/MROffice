/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A binary file of the Code workspace: a figure, a PDF or a data file that a
 * LaTeX project includes. Like every workspace file it is a code page, with
 * its path in content_json.codeWorkspace, but it holds no source: its bytes
 * are an attachment of that page, which content_json.binaryAsset names along
 * with the media type, size and SHA-256 they were uploaded with.
 *
 * Its Code window shows it read-only (CodeBinaryFilePreview), never in the
 * editor, and a LaTeX compile sends it by attachment id (latex-project.ts).
 * The API checks the attachment belongs to a page of the compiled page's
 * notebook that the member may read before it reads the bytes.
 *
 * The Explorer's Upload keeps text as text: a file that decodes as text
 * becomes an ordinary, editable code page, and only the rest becomes a binary
 * file. Images and PDFs are binary whatever their bytes decode as.
 */
import { apiFetchRaw, apiGet } from "@mrnote/core/web/lib/api.ts";
import { blobSha256Hex } from "@mrnote/core/web/lib/blob-sha256.ts";
import { extensionFromCodeFilename } from "./code-file-import";
import { sha256Hex } from "./execution-artifact";
import { isLatexProjectAssetId } from "./latex-project";
import { uploadPageAttachment } from "@mrnote/core/web/lib/page-attachment-upload.ts";

export const WORKSPACE_BINARY_ASSET_KEY = "binaryAsset";
/** A binary file is never edited or run; its page keeps a neutral language. */
export const WORKSPACE_BINARY_LANGUAGE = "plaintext";
/**
 * The largest file the Code window fetches to show it (an image drawn, a
 * PDF's pages); a larger one shows its details and downloads from storage.
 * A binary file itself has no size of its own: the workspace's storage
 * quota bounds it, as it bounds every upload, and a LaTeX compile that would
 * send more than one project may says so when it is run (latex-project.ts).
 */
export const MAX_WORKSPACE_BINARY_PREVIEW_BYTES = 64 * 1024 * 1024;

export interface WorkspaceBinaryAsset {
  readonly attachmentId: string;
  readonly mediaType: string;
  readonly sizeBytes: number;
  /** Lowercase hex SHA-256 of the uploaded bytes. */
  readonly sha256: string;
}

export type WorkspaceBinaryPreviewKind = "image" | "pdf" | "other";

/** binaryAsset's keys, sorted: the object carries exactly these. */
const ASSET_KEYS = ["attachmentId", "mediaType", "sha256", "sizeBytes"];
const MEDIA_TYPE =
  /^[a-z0-9][a-z0-9!#$&^_.+-]{0,63}\/[a-z0-9][a-z0-9!#$&^_.+-]{0,126}$/;
const SHA256 = /^[0-9a-f]{64}$/;
const OCTET_STREAM = "application/octet-stream";

/** Formats that are binary whatever their bytes happen to decode as. */
const MEDIA_TYPE_BY_EXTENSION: Readonly<Record<string, string>> = {
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".webp": "image/webp",
  ".bmp": "image/bmp",
  ".tif": "image/tiff",
  ".tiff": "image/tiff",
  ".pdf": "application/pdf",
};
/** What the preview draws as an image; anything else shows its details. */
const PREVIEW_IMAGE_TYPES: ReadonlySet<string> = new Set([
  "image/png",
  "image/jpeg",
  "image/gif",
  "image/webp",
]);

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function knownMediaType(filename: string): string | null {
  const extension = extensionFromCodeFilename(filename)?.toLowerCase();
  return extension &&
    Object.prototype.hasOwnProperty.call(MEDIA_TYPE_BY_EXTENSION, extension)
    ? MEDIA_TYPE_BY_EXTENSION[extension]
    : null;
}

/**
 * The binary asset a code page's content names, or null when it names none
 * or names one malformed. A page that holds source is a text file whatever
 * else it carries, so a binary file can never shadow typed text.
 */
export function parseWorkspaceBinaryAsset(
  contentJson: unknown,
): WorkspaceBinaryAsset | null {
  if (!isRecord(contentJson)) return null;
  if (contentJson.content !== undefined && contentJson.content !== "") {
    return null;
  }
  const raw = contentJson[WORKSPACE_BINARY_ASSET_KEY];
  if (!isRecord(raw)) return null;
  const keys = Object.keys(raw).sort();
  if (
    keys.length !== ASSET_KEYS.length ||
    keys.some((key, index) => key !== ASSET_KEYS[index])
  ) {
    return null;
  }
  const { attachmentId, mediaType, sizeBytes, sha256 } = raw;
  if (
    typeof attachmentId !== "string" ||
    !isLatexProjectAssetId(attachmentId) ||
    typeof mediaType !== "string" ||
    !MEDIA_TYPE.test(mediaType) ||
    typeof sha256 !== "string" ||
    !SHA256.test(sha256) ||
    typeof sizeBytes !== "number" ||
    !Number.isSafeInteger(sizeBytes) ||
    sizeBytes < 0
  ) {
    return null;
  }
  return { attachmentId, mediaType, sizeBytes, sha256 };
}

/** A page's content with its uploaded asset named in it. */
export function withWorkspaceBinaryAsset(
  contentJson: Record<string, unknown>,
  asset: WorkspaceBinaryAsset,
): Record<string, unknown> {
  return { ...contentJson, content: "", [WORKSPACE_BINARY_ASSET_KEY]: asset };
}

/** Whether a file of this name is uploaded as binary without reading it as text. */
export function isWorkspaceBinaryFilename(filename: string): boolean {
  return knownMediaType(filename) !== null;
}

export function workspaceBinaryPreviewKind(
  asset: Pick<WorkspaceBinaryAsset, "mediaType">,
): WorkspaceBinaryPreviewKind {
  if (PREVIEW_IMAGE_TYPES.has(asset.mediaType)) return "image";
  return asset.mediaType === "application/pdf" ? "pdf" : "other";
}

/**
 * A known format's own type first (the API refuses a .png whose bytes are not
 * PNG), then what the API or the browser declared, else octet-stream.
 */
function workspaceBinaryMediaType(
  filename: string,
  ...declared: readonly unknown[]
): string {
  const known = knownMediaType(filename);
  if (known) return known;
  for (const candidate of declared) {
    if (typeof candidate !== "string") continue;
    const mediaType = candidate.split(";", 1)[0].trim().toLowerCase();
    if (MEDIA_TYPE.test(mediaType) && mediaType !== OCTET_STREAM) {
      return mediaType;
    }
  }
  return OCTET_STREAM;
}

interface AttachmentUploadReceipt {
  readonly attachment_id?: unknown;
  readonly mime_type?: unknown;
  readonly size_bytes?: unknown;
}

/**
 * Upload `file` as an attachment of its binary file's page and describe it:
 * in one request, or in parts past 32 MiB (lib/page-attachment-upload). Its
 * digest is taken from the same File, read in slices past 64 MiB, so the
 * file is never held whole.
 */
export async function uploadWorkspaceBinaryAsset(
  pageId: string,
  file: File,
  options: { readonly notebookId?: string | null } = {},
): Promise<WorkspaceBinaryAsset> {
  const [receipt, sha256] = await Promise.all([
    uploadPageAttachment(pageId, file, {
      title: file.name,
      notebookId: options.notebookId,
    }) as Promise<AttachmentUploadReceipt | null>,
    blobSha256Hex(file),
  ]);
  const attachmentId = receipt?.attachment_id;
  if (
    typeof attachmentId !== "string" ||
    !isLatexProjectAssetId(attachmentId) ||
    receipt?.size_bytes !== file.size
  ) {
    throw new Error("The upload receipt does not describe the file.");
  }
  return {
    attachmentId,
    mediaType: workspaceBinaryMediaType(file.name, receipt.mime_type, file.type),
    sizeBytes: file.size,
    sha256,
  };
}

/** A short-lived link to the stored bytes, for a download of a large file. */
export async function workspaceBinaryAssetDownloadUrl(
  asset: Pick<WorkspaceBinaryAsset, "attachmentId">,
): Promise<string | null> {
  const answer = await apiGet<{ url?: unknown }>(
    `/api/v1/attachments/${encodeURIComponent(asset.attachmentId)}/url`,
  );
  return typeof answer?.url === "string" && /^https?:\/\//i.test(answer.url)
    ? answer.url
    : null;
}

/** The asset's bytes, refused unless they are the bytes that were uploaded. */
export async function readWorkspaceBinaryAsset(
  asset: WorkspaceBinaryAsset,
  signal?: AbortSignal,
): Promise<Uint8Array<ArrayBuffer>> {
  const response = await apiFetchRaw(
    `/api/v1/attachments/${encodeURIComponent(asset.attachmentId)}/download`,
    { method: "GET", signal },
    { requireCsrf: false, contentType: "" },
  );
  const bytes = new Uint8Array(await response.arrayBuffer());
  if (
    bytes.byteLength !== asset.sizeBytes ||
    (await sha256Hex(bytes)) !== asset.sha256
  ) {
    throw new Error("The file's bytes are not the ones that were uploaded.");
  }
  return bytes;
}

/** "640 kB", "3.2 MB": decimal units, as the LaTeX project budget is stated. */
export function formatWorkspaceBinarySize(
  bytes: number,
  locale?: string,
): string {
  const [unit, value] =
    bytes < 1_000
      ? (["byte", bytes] as const)
      : bytes < 1_000_000
        ? (["kilobyte", bytes / 1_000] as const)
        : (["megabyte", bytes / 1_000_000] as const);
  const options: Intl.NumberFormatOptions = {
    style: "unit",
    unit,
    unitDisplay: "short",
    maximumFractionDigits: unit === "byte" ? 0 : 1,
  };
  try {
    return new Intl.NumberFormat(locale, options).format(value);
  } catch {
    return new Intl.NumberFormat(undefined, options).format(value);
  }
}
