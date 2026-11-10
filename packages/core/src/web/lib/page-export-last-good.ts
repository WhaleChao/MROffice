/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { downloadUrlFile } from "./download-file";
import type { PageExportLastGoodCopy } from "./page-export";
import { getResourceVersionDownload } from "./resources-sdk";

/** The stored file to offer when an export cannot be built or is not the one
 * wanted: the last export (the file as last saved), else the original upload. */
export function preferredLastGoodCopy(copies: readonly PageExportLastGoodCopy[]): PageExportLastGoodCopy | null {
  return copies.find((copy) => copy.kind === "last_export") ?? copies.find((copy) => copy.kind === "original") ?? null;
}

/** Fetch the stored file's short-lived URL and download it. */
export async function downloadLastGoodCopy(
  copy: PageExportLastGoodCopy,
  fileName: string,
  extension = "xlsx",
): Promise<boolean> {
  try {
    const answer = await getResourceVersionDownload(copy.resourceId, copy.versionId);
    return downloadUrlFile(`${fileName}.${extension}`, answer.download_url);
  } catch {
    return false;
  }
}
