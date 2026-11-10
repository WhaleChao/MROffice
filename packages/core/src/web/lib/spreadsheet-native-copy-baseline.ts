/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { apiGet, apiPost } from "./api";
import { getPage } from "./notebook-sdk";
import { getWorkbookCollabV2 } from "./page-collab-sdk";
import { pageDocumentJsonEquals as equal } from "./page-save-acknowledgement";

type Json = Record<string, unknown>;
type Page = Awaited<ReturnType<typeof getPage>>;
const record = (value: unknown): Json | null =>
  value !== null && typeof value === "object" && !Array.isArray(value)
    ? (value as Json)
    : null;
const digest = (value: unknown): value is string =>
  typeof value === "string" && /^[0-9a-f]{64}$/.test(value);

export async function getSpreadsheetNativeCopyCapabilities(
  pageId: string,
  version: number,
  signal?: AbortSignal,
) {
  const value = record(
    await apiGet<unknown>(
      `/api/v1/pages/${encodeURIComponent(pageId)}/native-xlsx-copy-capabilities`,
      { signal },
    ),
  );
  const matched =
    value?.page_id === pageId && value?.content_version === version;
  const available = matched && value?.worksheet_source === true;
  return {
    available,
    baselinePrepareAvailable:
      matched &&
      !available &&
      value?.baseline_prepare_available === true &&
      digest(value?.content_sha256),
    contentSha256:
      matched && digest(value?.content_sha256) ? value.content_sha256 : null,
  };
}

/** Explicit Copy only. Eligibility is not source authority: the existing
 * prepare/publish endpoints revalidate version, body, actor and absence of a
 * source. No page-open/hover caller may invoke this function. */
export async function ensureSpreadsheetNativeCopySource(
  source: Page,
  actorIncarnation: string,
  isCurrent: () => boolean,
): Promise<boolean> {
  if (
    !isCurrent() ||
    source.page_type !== "spreadsheet" ||
    !Number.isSafeInteger(source.content_version) ||
    !source.content_json
  )
    return false;
  const version = source.content_version!;
  const capability = await getSpreadsheetNativeCopyCapabilities(
    source.id,
    version,
  );
  if (!isCurrent()) return false;
  if (capability.available) return true;
  if (
    !capability.baselinePrepareAvailable ||
    capability.contentSha256 !== source.content_sha256
  )
    return false;
  const hash = capability.contentSha256;
  const current = async () => {
    if (!isCurrent()) return false;
    const actor = await getWorkbookCollabV2(source.id);
    if (!isCurrent() || actor.actorMembershipIncarnation !== actorIncarnation)
      return false;
    const page = await getPage(source.id);
    return (
      isCurrent() &&
      page.id === source.id &&
      page.page_type === "spreadsheet" &&
      page.content_version === version &&
      page.content_sha256 === hash &&
      equal(page.content_json, source.content_json)
    );
  };
  if (!(await current())) return false;
  const request = {
    expected_content_version: version,
    expected_content_sha256: hash,
  };
  const endpoint = `/api/v1/pages/${encodeURIComponent(source.id)}/native-xlsx-baseline`;
  const prepared = record(
    await apiPost<unknown>(`${endpoint}/prepare`, request),
  );
  if (
    !prepared ||
    prepared.page_id !== source.id ||
    prepared.content_version !== version ||
    prepared.content_sha256 !== hash ||
    prepared.prepared !== true ||
    prepared.published !== false ||
    !(await current())
  )
    return false;
  const published = record(
    await apiPost<unknown>(`${endpoint}/publish`, request),
  );
  if (
    !published ||
    published.page_id !== source.id ||
    published.content_version !== version ||
    published.content_sha256 !== hash ||
    published.published !== true ||
    published.page_changed !== false ||
    !digest(published.package_sha256) ||
    typeof published.resource_id !== "string" ||
    !published.resource_id ||
    typeof published.resource_version_id !== "string" ||
    !published.resource_version_id ||
    !(await current())
  )
    return false;
  const actual = await getSpreadsheetNativeCopyCapabilities(source.id, version);
  return isCurrent() && actual.available;
}
