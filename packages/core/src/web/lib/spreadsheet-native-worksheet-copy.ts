/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { apiGet, apiPost } from "./api";
import { spreadsheetPageContentSchema } from "./native-spreadsheet-schemas";
import { pageDocumentJsonEquals as equal } from "./page-save-acknowledgement";
import type { SpreadsheetNativeCopySavedReceipt } from "./spreadsheet-native-copy";
import type { WorkbookProjectionOperationReceipt } from "./page-collab-sdk";
import type { PageSaveAcknowledgement } from "../hooks/useAutosavedPageDocument";

export type SpreadsheetCopyCanonicalAcknowledgement =
  PageSaveAcknowledgement & {
    workbookCanonicalReceipt?: WorkbookProjectionOperationReceipt;
  };

type Json = Record<string, unknown>;
const record = (value: unknown): Json | null =>
  value !== null && typeof value === "object" && !Array.isArray(value)
    ? (value as Json)
    : null;
const digest = (value: unknown): value is string =>
  typeof value === "string" && /^[0-9a-f]{64}$/.test(value);
const ids = (value: unknown): value is string[] =>
  Array.isArray(value) &&
  value.every((id) => typeof id === "string" && id.length > 0) &&
  new Set(value).size === value.length;

export async function hasSpreadsheetNativeWorksheetCopySource(
  pageId: string,
  version: number,
  signal: AbortSignal,
): Promise<boolean> {
  const value = record(
    await apiGet<unknown>(
      `/api/v1/pages/${encodeURIComponent(pageId)}/native-xlsx-copy-capabilities`,
      { signal },
    ),
  );
  return Boolean(
    value &&
    value.page_id === pageId &&
    value.content_version === version &&
    value.worksheet_source === true,
  );
}

export interface SpreadsheetNativeWorksheetCopyRequest {
  pageId: string;
  contentVersion: number;
  sourceSheetId: string;
  copiedName: string;
  source: Json;
  draft: Json;
}

export interface SpreadsheetNativeWorksheetCopyPlan {
  content: Json;
  receipt: Json;
  insertedSheetIds: string[];
  insertedObjectIds: string[];
}

export function readSpreadsheetNativeWorksheetCopyPlan(
  value: unknown,
  request: SpreadsheetNativeWorksheetCopyRequest,
): SpreadsheetNativeWorksheetCopyPlan | null {
  const response = record(value),
    receipt = record(response?.receipt),
    content = record(response?.content);
  const sheetIds = response?.inserted_sheet_ids,
    objectIds = response?.inserted_object_ids;
  const objectMap = record(receipt?.modelObjectIdMap);
  if (
    !response ||
    !receipt ||
    !content ||
    !objectMap ||
    !ids(sheetIds) ||
    sheetIds.length !== 1 ||
    !ids(objectIds) ||
    response.plan_id !== receipt.planId ||
    typeof receipt.planId !== "string" ||
    !receipt.planId ||
    response.expected_content_version !== request.contentVersion ||
    receipt.schemaVersion !== "xlsx-native-worksheet-copy-v1" ||
    receipt.pageId !== request.pageId ||
    receipt.sourceContentVersion !== request.contentVersion ||
    receipt.sourceSheetId !== request.sourceSheetId ||
    receipt.copiedSheetId !== sheetIds[0] ||
    receipt.copiedSheetName !== request.copiedName ||
    typeof receipt.sourceResourceId !== "string" ||
    !receipt.sourceResourceId ||
    !digest(receipt.sourcePackageSha256) ||
    !digest(receipt.outputPackageSha256) ||
    !digest(receipt.sourceContentSha256) ||
    !ids(Object.values(objectMap)) ||
    objectIds.some((id) => !Object.values(objectMap).includes(id)) ||
    !equal(content.nativeWorksheetCopies, [receipt]) ||
    !spreadsheetPageContentSchema.safeParse(content).success
  )
    return null;
  const sourceSheets = request.source.sheets;
  const sheets = content.sheets;
  if (
    !Array.isArray(sourceSheets) ||
    !Array.isArray(sheets) ||
    sheets.length !== sourceSheets.length + 1
  )
    return null;
  const sourceIndex = sourceSheets.findIndex(
    (sheet) => record(sheet)?.id === request.sourceSheetId,
  );
  const sourceSheet = record(sourceSheets[sourceIndex]);
  if (
    sourceIndex < 0 ||
    !sourceSheet ||
    sourceSheet.name !== receipt.sourceSheetName ||
    !equal(
      sheets.filter((sheet) => record(sheet)?.id !== sheetIds[0]),
      sourceSheets,
    ) ||
    record(sheets[sourceIndex + 1])?.id !== sheetIds[0] ||
    sourceSheets.some((sheet) => record(sheet)?.id === sheetIds[0])
  )
    return null;
  const copied = record(sheets[sourceIndex + 1])!;
  const copiedObjects = Array.isArray(copied.objects) ? copied.objects : [];
  // Reviewed deletion may remove a copied object; any remaining object must
  // still be one of the server-materialized independent identities.
  if (
    !equal(
      copiedObjects.map((obj) => record(obj)?.id),
      objectIds,
    )
  )
    return null;
  return {
    content,
    receipt,
    insertedSheetIds: sheetIds,
    insertedObjectIds: objectIds,
  };
}

export async function requestSpreadsheetNativeWorksheetCopy(
  request: SpreadsheetNativeWorksheetCopyRequest,
): Promise<SpreadsheetNativeWorksheetCopyPlan> {
  const result = await apiPost<unknown>(
    `/api/v1/pages/${encodeURIComponent(request.pageId)}/native-xlsx-copy-plan`,
    {
      expected_content_version: request.contentVersion,
      source_sheet_id: request.sourceSheetId,
      copied_name: request.copiedName,
      draft_content: request.draft,
    },
  );
  const plan = readSpreadsheetNativeWorksheetCopyPlan(result, request);
  if (!plan) throw new Error("native_xlsx_copy_response_mismatch");
  return plan;
}

function authoredContent(value: Json): Json {
  const content = structuredClone(value);
  delete content.activeSheetId;
  for (const sheet of Array.isArray(content.sheets) ? content.sheets : []) {
    const cells = record(record(sheet)?.cells);
    for (const cell of Object.values(cells ?? {})) {
      const entry = record(cell);
      if (entry) {
        if (typeof entry.display === "string") delete entry.display;
        delete entry.chartValueCache;
      }
    }
  }
  return content;
}

/** Only the normal server save acknowledgement can close the private preview.
 * A successful plan, a local V2 update, or a generic HTTP 200 is insufficient.
 */
export function readSpreadsheetNativeWorksheetCopySavedReceipt(
  value: unknown,
  request: SpreadsheetNativeWorksheetCopyRequest,
  plan: SpreadsheetNativeWorksheetCopyPlan,
  actorIncarnation: string,
  acknowledgement?: SpreadsheetCopyCanonicalAcknowledgement | null,
): SpreadsheetNativeCopySavedReceipt | null {
  const saved = record(value),
    content = record(saved?.content_json);
  const status = record(saved?.ai_status_json);
  const marker = record(
    saved?.editable_page_patchback ?? status?.editable_page_patchback,
  );
  const copy = record(marker?.native_worksheet_copy);
  if (
    !saved ||
    !content ||
    !marker ||
    !copy ||
    !actorIncarnation ||
    saved.id !== request.pageId ||
    saved.content_version !== request.contentVersion + 1 ||
    marker.status !== "succeeded" ||
    marker.adapter !== "xlsx" ||
    marker.loss_entry_count !== 0 ||
    marker.resource_id !== plan.receipt.sourceResourceId ||
    typeof marker.resource_version_id !== "string" ||
    !marker.resource_version_id ||
    !digest(marker.output_sha256) ||
    copy.inserted !== true ||
    copy.saved_page_id !== request.pageId ||
    copy.saved_content_version !== saved.content_version ||
    !equal(copy.receipt, plan.receipt) ||
    !acknowledgedWorksheetContent(content, request, plan, acknowledgement)
  )
    return null;
  return {
    schemaVersion: "xlsx-native-copy-saved-v1",
    savedPageId: request.pageId,
    sourceContentVersion: request.contentVersion,
    savedVersion: request.contentVersion + 1,
    actorIncarnation,
    planId: String(plan.receipt.planId),
    insertedSheetIds: plan.insertedSheetIds,
    insertedObjectIds: plan.insertedObjectIds,
    sourcePreserved: true,
    native: {
      resourceId: marker.resource_id as string,
      resourceVersionId: marker.resource_version_id,
      sourcePackageSha256: String(plan.receipt.sourcePackageSha256),
      outputPackageSha256: marker.output_sha256,
      lossCount: 0,
    },
  };
}

function acknowledgedWorksheetContent(
  content: Json,
  request: SpreadsheetNativeWorksheetCopyRequest,
  plan: SpreadsheetNativeWorksheetCopyPlan,
  acknowledgement?: SpreadsheetCopyCanonicalAcknowledgement | null,
): boolean {
  const current = authoredContent(content),
    expected = authoredContent(plan.content);
  if (equal(current, expected)) return true;
  const context = acknowledgement?.workbookCanonicalReceipt;
  const witness = context?.nativeChartCacheAck;
  if (
    !acknowledgement ||
    !record(acknowledgement.page) ||
    !record(acknowledgement.draft) ||
    acknowledgement.source !== "patch" ||
    acknowledgement.requestBaseContentVersion !== request.contentVersion ||
    acknowledgement.page.id !== request.pageId ||
    acknowledgement.page.content_version !== request.contentVersion + 1 ||
    !acknowledgement.draft.content_json ||
    !equal(authoredContent(acknowledgement.draft.content_json), expected) ||
    !context ||
    !record(context.contentJson) ||
    context.pageId !== request.pageId ||
    context.contentVersion !== request.contentVersion + 1 ||
    typeof context.operationId !== "string" ||
    !context.operationId ||
    typeof context.actorMembershipIncarnation !== "string" ||
    !context.actorMembershipIncarnation ||
    !digest(context.sharedContentSha256) ||
    !Number.isSafeInteger(context.workbookRevision) ||
    context.workbookRevision < 0 ||
    !Number.isSafeInteger(context.canonicalWorkbookRevision) ||
    context.canonicalWorkbookRevision! <= context.workbookRevision ||
    !witness ||
    witness.schemaVersion !== "workbook-native-chart-cache-ack-v1" ||
    witness.sourceOperationId !== context.operationId ||
    typeof witness.derivedOperationId !== "string" ||
    !witness.derivedOperationId ||
    witness.derivedOperationId === context.operationId ||
    !Array.isArray(witness.entries) ||
    witness.entries.length === 0 ||
    witness.entries.length > 1024 ||
    !equal(authoredContent(context.contentJson), current) ||
    !equal(authoredContent(acknowledgement.page.content_json ?? {}), current)
  )
    return false;
  const seen = new Set<string>();
  for (const entry of witness.entries) {
    if (!record(entry)) return false;
    if (
      !plan.insertedSheetIds.includes(entry.sheetId) ||
      !plan.insertedObjectIds.includes(entry.objectId) ||
      !Object.values(record(plan.receipt.partMap) ?? {}).includes(entry.part) ||
      !Object.values(record(plan.receipt.modelObjectIdMap) ?? {}).includes(
        entry.objectId,
      )
    )
      return false;
    const key = JSON.stringify([entry.sheetId, entry.objectId]);
    if (seen.has(key)) return false;
    seen.add(key);
    const find = (document: Json) => {
      const sheets = (Array.isArray(document.sheets) ? document.sheets : [])
        .map(record)
        .filter((sheet) => sheet?.id === entry.sheetId);
      if (sheets.length !== 1) return null;
      const objects = (
        Array.isArray(sheets[0]?.objects) ? sheets[0].objects : []
      )
        .map(record)
        .filter((obj) => obj?.id === entry.objectId);
      return objects.length === 1 ? objects[0] : null;
    };
    const before = find(expected),
      after = find(current);
    if (
      !before ||
      !after ||
      before.type !== "chart" ||
      after.type !== "chart" ||
      Object.hasOwn(before, "data") ||
      !isChartData(after.data) ||
      record(before.source)?.part !== entry.part ||
      record(after.source)?.part !== entry.part ||
      !equal(before.source, after.source) ||
      !equal(before.ranges, after.ranges) ||
      !equal(before.seriesRefs, after.seriesRefs)
    )
      return false;
    delete after.data;
  }
  return equal(current, expected);
}

function isChartData(value: unknown): boolean {
  const data = record(value);
  return Boolean(
    data &&
    Object.keys(data).every((key) => key === "headers" || key === "rows") &&
    Array.isArray(data.headers) &&
    data.headers.every((value) => typeof value === "string") &&
    Array.isArray(data.rows) &&
    data.rows.every(
      (row) =>
        Array.isArray(row) && row.every((value) => typeof value === "string"),
    ),
  );
}
