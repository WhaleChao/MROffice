/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { apiPost } from "./api";
import { getPage } from "./notebook-sdk";
import { getWorkbookCollabV2 } from "./page-collab-sdk";
import { pageDocumentJsonEquals as equal } from "./page-save-acknowledgement";
import { spreadsheetPageContentSchema } from "./native-spreadsheet-schemas";
import { spreadsheetPersistenceJson } from "./spreadsheet/persistence-json";
import type { CopilotSpreadsheetContent } from "../__cut__/mfc2dd2a6e6";
import { withCurrentSpreadsheetCopyView } from "./spreadsheet-native-copy";
import {
  captureSavedPageHotCacheScope,
  isSavedPageHotCacheScopeCurrent,
} from "./saved-page-hot-cache";

type Json = Record<string, unknown>;
const record = (value: unknown): Json | null =>
  value !== null && typeof value === "object" && !Array.isArray(value)
    ? (value as Json)
    : null;

export interface SpreadsheetNativeCopyHistoryStep {
  pageId: string;
  planId: string;
  resourceId: string;
  copyExportVersionId: string;
  sourceSheetId: string;
  copiedSheetId: string;
  actorIncarnation: string;
  actorId: string;
}

function authoring(content: Json) {
  const result = structuredClone(content);
  for (const key of [
    "activeSheetId",
    "nativeWorksheetCopies",
    "nativeWorksheetCopyHistory",
  ])
    delete result[key];
  for (const sheet of Array.isArray(result.sheets) ? result.sheets : []) {
    for (const cell of Object.values(record(record(sheet)?.cells) ?? {})) {
      const value = record(cell);
      if (value) {
        if (typeof value.display === "string") delete value.display;
        delete value.chartValueCache;
      }
    }
  }
  return result;
}

export function readSpreadsheetNativeCopyHistoryPlan(
  value: unknown,
  step: SpreadsheetNativeCopyHistoryStep,
  operation: "undo" | "redo",
  version: number,
  source: Json,
): { content: Json; receipt: Json } | null {
  const plan = record(value),
    content = record(plan?.content),
    receipt = record(plan?.receipt);
  if (
    !content ||
    !receipt ||
    !spreadsheetPageContentSchema.safeParse(content).success ||
    receipt.schemaVersion !== "xlsx-native-worksheet-copy-history-v1" ||
    receipt.operation !== operation ||
    receipt.pageId !== step.pageId ||
    receipt.sourceContentVersion !== version ||
    receipt.copyExportVersionId !== step.copyExportVersionId ||
    receipt.copyPlanId !== step.planId ||
    typeof receipt.sourceContentSha256 !== "string" ||
    !/^[0-9a-f]{64}$/.test(receipt.sourceContentSha256) ||
    typeof receipt.templatePackageSha256 !== "string" ||
    !/^[0-9a-f]{64}$/.test(receipt.templatePackageSha256) ||
    !["original", "export"].includes(
      String(receipt.templateResourceVersionKind),
    ) ||
    !(
      receipt.templateResourceVersionId === null ||
      typeof receipt.templateResourceVersionId === "string"
    ) ||
    !equal(content.nativeWorksheetCopyHistory, receipt)
  )
    return null;
  const previous = source.sheets,
    next = content.sheets;
  if (!Array.isArray(previous) || !Array.isArray(next)) return null;
  const withoutCopy = (sheets: unknown[]) =>
    sheets.filter((sheet) => record(sheet)?.id !== step.copiedSheetId);
  if (operation === "undo") {
    if (
      previous.filter((s) => record(s)?.id === step.copiedSheetId).length !==
        1 ||
      next.some((s) => record(s)?.id === step.copiedSheetId) ||
      !equal(
        authoring({ sheets: withoutCopy(previous) }),
        authoring({ sheets: next }),
      )
    )
      return null;
  } else {
    const sourceIndex = next.findIndex(
      (s) => record(s)?.id === step.sourceSheetId,
    );
    if (
      previous.some((s) => record(s)?.id === step.copiedSheetId) ||
      next.filter((s) => record(s)?.id === step.copiedSheetId).length !== 1 ||
      sourceIndex < 0 ||
      record(next[sourceIndex + 1])?.id !== step.copiedSheetId ||
      !equal(
        authoring({ sheets: withoutCopy(next) }),
        authoring({ sheets: previous }),
      )
    )
      return null;
  }
  const oldRoot = authoring({ ...source, sheets: [] }),
    newRoot = authoring({ ...content, sheets: [] });
  return equal(oldRoot, newRoot) ? { content, receipt } : null;
}

export function spreadsheetNativeCopyHistorySaved(
  value: unknown,
  step: SpreadsheetNativeCopyHistoryStep,
  version: number,
  plan: { content: Json; receipt: Json },
): boolean {
  const page = record(value),
    content = record(page?.content_json);
  const marker = record(
    page?.editable_page_patchback ??
      record(page?.ai_status_json)?.editable_page_patchback,
  );
  const copy = record(marker?.native_worksheet_copy);
  return Boolean(
    page &&
    content &&
    marker &&
    copy &&
    page.id === step.pageId &&
    page.content_version === version + 1 &&
    marker.status === "succeeded" &&
    marker.adapter === "xlsx" &&
    marker.loss_entry_count === 0 &&
    marker.resource_id === step.resourceId &&
    typeof marker.resource_version_id === "string" &&
    marker.resource_version_id &&
    typeof marker.output_sha256 === "string" &&
    /^[0-9a-f]{64}$/.test(marker.output_sha256) &&
    copy.saved_page_id === step.pageId &&
    copy.saved_content_version === version + 1 &&
    equal(copy.history, plan.receipt) &&
    equal(authoring(content), authoring(plan.content)),
  );
}

export async function applySpreadsheetNativeCopyHistory(
  step: SpreadsheetNativeCopyHistoryStep,
  operation: "undo" | "redo",
  host: {
    isCurrent(): boolean;
    readSource(): CopilotSpreadsheetContent;
    normalizeSavedContent(content: Json): CopilotSpreadsheetContent;
    flush(): Promise<unknown>;
    save(content: Json): Promise<unknown>;
    accept(content: Json): boolean;
  },
): Promise<boolean> {
  const actorScope = captureSavedPageHotCacheScope();
  const current = () =>
    host.isCurrent() && isSavedPageHotCacheScopeCurrent(actorScope);
  if (!current() || !(await host.flush()) || !current()) return false;
  const source = spreadsheetPersistenceJson(host.readSource());
  const unchanged = () =>
    current() && equal(spreadsheetPersistenceJson(host.readSource()), source);
  const page = await getPage(step.pageId);
  if (
    !page.content_json ||
    !Number.isSafeInteger(page.content_version) ||
    !unchanged() ||
    !equal(
      spreadsheetPersistenceJson(host.normalizeSavedContent(page.content_json)),
      source,
    )
  )
    return false;
  const actor = await getWorkbookCollabV2(step.pageId);
  if (
    !unchanged() ||
    actor.actorMembershipIncarnation !== step.actorIncarnation
  )
    return false;
  const version = page.content_version!;
  const response = await apiPost<unknown>(
    `/api/v1/pages/${encodeURIComponent(step.pageId)}/native-xlsx-copy-history-plan`,
    {
      expected_content_version: version,
      copy_export_version_id: step.copyExportVersionId,
      operation,
    },
  );
  const plan = readSpreadsheetNativeCopyHistoryPlan(
    response,
    step,
    operation,
    version,
    page.content_json,
  );
  if (!plan || !unchanged()) return false;
  const saved = await host.save(withCurrentSpreadsheetCopyView(plan.content, source));
  if (!current()) return false;
  if (!spreadsheetNativeCopyHistorySaved(saved, step, version, plan)) {
    if (
      !saved ||
      !spreadsheetNativeCopyHistorySaved(
        await getPage(step.pageId),
        step,
        version,
        plan,
      ) ||
      !current()
    )
      return false;
  }
  return host.accept(plan.content);
}
