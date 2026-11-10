/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { z } from "zod";
import { apiPost } from "./api";
import { spreadsheetPageContentSchema } from "./native-spreadsheet-schemas";
import { pageDocumentJsonEquals as equal } from "./page-save-acknowledgement";
import { spreadsheetPersistenceJson } from "./spreadsheet/persistence-json";
import type { SpreadsheetNativeCopySavedReceipt } from "./spreadsheet-native-copy";

type Json = Record<string, unknown>;
const record = (value: unknown): Json | null =>
  value !== null && typeof value === "object" && !Array.isArray(value)
    ? (value as Json)
    : null;
const identifier = z.string().min(1);
const version = z.number().int().positive().max(Number.MAX_SAFE_INTEGER);
const digest = z.string().regex(/^[0-9a-f]{64}$/);
const nonVisualId = z
  .string()
  .regex(/^(0|[1-9]\d{0,9})$/)
  .refine((value) => Number(value) <= 4_294_967_295);
const part = z
  .string()
  .refine(
    (value) =>
      value.startsWith("xl/") &&
      !/[\\\u0000-\u001f?#]/.test(value) &&
      value
        .split("/")
        .every(
          (segment) => segment !== "" && segment !== "." && segment !== "..",
        ),
  );
const uniqueValues = (mapping: Record<string, string>) =>
  new Set(Object.values(mapping)).size === Object.keys(mapping).length;
const partMap = z.record(part, part).refine(uniqueValues);
const relationshipMap = z.record(identifier, identifier).refine(uniqueValues);
const receiptSchema = z
  .object({
    schemaVersion: z.literal("xlsx-native-object-copy-v1"),
    planId: identifier,
    pageId: identifier,
    sourceContentVersion: version,
    sourceContentSha256: digest,
    sourcePackageSha256: digest,
    outputPackageSha256: digest,
    sourceResourceId: identifier,
    sourceResourceVersionId: identifier.nullable(),
    sourceResourceVersionKind: z.enum(["original", "export"]),
    sheetId: identifier,
    sourceObjectId: identifier,
    copiedObjectId: identifier,
    objectKind: z.enum(["shape", "image", "chart"]),
    worksheetPart: part,
    drawingPart: part,
    sourceNonVisualId: nonVisualId,
    copiedNonVisualId: nonVisualId,
    relationshipIdMap: relationshipMap,
    partMap,
  })
  .strict()
  .refine(
    (value) =>
      value.sourceObjectId !== value.copiedObjectId &&
      value.sourceNonVisualId !== value.copiedNonVisualId &&
      Object.entries(value.partMap).every(
        ([before, after]) => before !== after,
      ) &&
      Object.entries(value.relationshipIdMap).every(
        ([before, after]) => before !== after,
      ) &&
      (value.sourceResourceVersionKind !== "export" ||
        value.sourceResourceVersionId !== null) &&
      (value.objectKind === "shape" ||
        (Object.keys(value.partMap).length > 0 &&
          Object.keys(value.relationshipIdMap).length > 0)),
  );
const historySchema = z
  .object({
    schemaVersion: z.literal("xlsx-native-object-copy-history-v1"),
    operation: z.enum(["undo", "redo"]),
    pageId: identifier,
    sourceContentVersion: version,
    sourceContentSha256: digest,
    copyExportVersionId: identifier,
    copyPlanId: identifier,
    insertedPackageSha256: digest,
    outputPackageSha256: digest,
    copiedNonVisualId: nonVisualId,
    partMap,
    relationshipIdMap: relationshipMap,
  })
  .strict();

export type SpreadsheetNativeObjectCopyReceipt = z.infer<typeof receiptSchema>;
export interface SpreadsheetNativeObjectCopyRequest {
  pageId: string;
  contentVersion: number;
  sheetId: string;
  sourceObjectId: string;
  source: Json;
  draft?: Json;
}
export interface SpreadsheetNativeObjectCopyPlan {
  content: Json;
  receipt: SpreadsheetNativeObjectCopyReceipt;
  insertedSheetIds: [];
  insertedObjectIds: [string];
}
export interface SpreadsheetNativeObjectCopyHistoryStep {
  receipt: SpreadsheetNativeObjectCopyReceipt;
  copyExportVersionId: string;
  actorIncarnation: string;
  actorId: string;
  /** Only the inserted object, not a whole-book undo snapshot. */
  copiedObject: Json;
}
export interface SpreadsheetNativeObjectCopyHistoryPlan {
  content: Json;
  receipt: z.infer<typeof historySchema>;
}

const controls = [
  "nativeWorksheetCopies",
  "nativeWorksheetCopyHistory",
  "nativeObjectCopies",
  "nativeObjectCopyHistory",
] as const;
function authoring(value: Json, stripControls = true): Json {
  const content = structuredClone(spreadsheetPersistenceJson(value));
  delete content.activeSheetId;
  if (stripControls) for (const key of controls) delete content[key];
  for (const sheet of Array.isArray(content.sheets) ? content.sheets : []) {
    for (const cell of Object.values(record(record(sheet)?.cells) ?? {})) {
      const entry = record(cell);
      if (entry) {
        if (typeof entry.display === "string") delete entry.display;
        delete entry.chartValueCache;
      }
    }
  }
  return content;
}
function sheetOf(content: Json, id: string): Json | null {
  if (!Array.isArray(content.sheets)) return null;
  const sheets = content.sheets.map(record);
  if (
    sheets.some((sheet) => !sheet || typeof sheet.id !== "string") ||
    new Set(sheets.map((sheet) => sheet?.id)).size !== sheets.length
  )
    return null;
  const matches = sheets.filter((sheet) => sheet?.id === id);
  return matches.length === 1 ? matches[0] : null;
}
function memberOf(
  sheet: Json | null,
  id: string,
): { object: Json; recipe: boolean } | null {
  if (!sheet) return null;
  const members: { object: Json; recipe: boolean }[] = [];
  for (const key of ["objects", "charts"] as const) {
    if (sheet[key] !== undefined && !Array.isArray(sheet[key])) return null;
    for (const value of (sheet[key] ?? []) as unknown[]) {
      const object = record(value);
      if (!object) return null;
      if (object.id === id) members.push({ object, recipe: key === "charts" });
    }
  }
  return members.length === 1 ? members[0] : null;
}
function withoutCopy(
  content: Json,
  sheetId: string,
  copyId: string,
  baseline: Json,
): Json | null {
  const result = authoring(content),
    sheet = sheetOf(result, sheetId);
  const selected = memberOf(sheet, copyId);
  if (!sheet || !selected || selected.recipe || !Array.isArray(sheet.objects))
    return null;
  const remaining = sheet.objects.filter(
    (value) => record(value)?.id !== copyId,
  );
  sheet.objects = remaining;
  if (remaining.length === 0 && !Object.hasOwn(baseline, "objects"))
    delete sheet.objects;
  return result;
}
function subset(expected: unknown, actual: unknown): boolean {
  const object = record(expected),
    target = record(actual);
  if (object)
    return Boolean(
      target &&
      Object.entries(object).every(
        ([key, value]) =>
          Object.hasOwn(target, key) && subset(value, target[key]),
      ),
    );
  if (Array.isArray(expected))
    return (
      Array.isArray(actual) &&
      expected.length === actual.length &&
      expected.every((value, index) => subset(value, actual[index]))
    );
  return equal(expected, actual);
}
function reviewedIntent(
  before: unknown,
  after: unknown,
  actual: unknown,
): boolean {
  const original = record(before),
    reviewed = record(after),
    output = record(actual);
  if (reviewed)
    return Boolean(
      output &&
      (!original ||
        Object.keys(original).every(
          (key) => Object.hasOwn(reviewed, key) || !Object.hasOwn(output, key),
        )) &&
      Object.entries(reviewed).every(
        ([key, value]) =>
          Object.hasOwn(output, key) &&
          reviewedIntent(original?.[key], value, output[key]),
      ),
    );
  if (Array.isArray(after))
    return (
      Array.isArray(actual) &&
      after.length === actual.length &&
      after.every((value, index) =>
        reviewedIntent(
          Array.isArray(before) ? before[index] : undefined,
          value,
          actual[index],
        ),
      )
    );
  return equal(after, actual);
}
function reviewedChartMirrors(
  before: Json,
  after: Json,
  kind: unknown,
): Json | null {
  const text = (value: unknown) =>
    typeof value === "string" ? value.trim() : "";
  const changed = ["title", "caption"].find(
    (key) => text(before[key]) !== text(after[key]),
  );
  const result = structuredClone(after);
  if (!changed) return result;
  const title = text(after[changed]),
    present = Boolean(title),
    caption = title || kind;
  for (const [key, expected] of [
    ["caption", caption],
    ["nativeTitlePresent", present],
  ] as const) {
    if (
      Object.hasOwn(after, key) &&
      !equal(before[key], after[key]) &&
      !equal(after[key], expected)
    )
      return null;
  }
  const oldLayout = record(record(before.chartFormat)?.layout) ?? {};
  const layout = record(record(result.chartFormat)?.layout);
  if (
    layout &&
    Object.hasOwn(layout, "title") &&
    !equal(layout.title, oldLayout.title) &&
    layout.title !== present
  )
    return null;
  result.title = caption;
  result.caption = caption;
  result.nativeTitlePresent = present;
  if (layout) layout.title = present;
  return result;
}
function boundCopy(
  object: Json,
  receipt: SpreadsheetNativeObjectCopyReceipt,
): boolean {
  const source = record(object.source);
  if (
    object.id !== receipt.copiedObjectId ||
    object.type !== receipt.objectKind ||
    !source ||
    source.drawingPart !== receipt.drawingPart
  )
    return false;
  if (receipt.objectKind === "shape")
    return source.nonVisualId === receipt.copiedNonVisualId;
  return (
    Object.values(receipt.partMap).includes(String(source.part)) &&
    Object.values(receipt.relationshipIdMap).includes(
      String(source.relationshipId),
    )
  );
}

/** Validate the server snapshot, never rewrite its native provenance. */
export function readSpreadsheetNativeObjectCopyPlan(
  value: unknown,
  request: SpreadsheetNativeObjectCopyRequest,
): SpreadsheetNativeObjectCopyPlan | null {
  try {
    const response = record(value),
      content = record(response?.content);
    const parsed = receiptSchema.safeParse(response?.receipt);
    if (
      !content ||
      !parsed.success ||
      !version.safeParse(request.contentVersion).success
    )
      return null;
    const receipt = parsed.data;
    if (
      receipt.pageId !== request.pageId ||
      receipt.sourceContentVersion !== request.contentVersion ||
      receipt.sheetId !== request.sheetId ||
      receipt.sourceObjectId !== request.sourceObjectId ||
      response?.plan_id !== receipt.planId ||
      response.expected_content_version !== request.contentVersion ||
      !equal(response.inserted_sheet_ids, []) ||
      !equal(response.inserted_object_ids, [receipt.copiedObjectId]) ||
      !equal(content.nativeObjectCopies, [receipt]) ||
      controls.some(
        (key) => key !== "nativeObjectCopies" && Object.hasOwn(content, key),
      ) ||
      !spreadsheetPageContentSchema.safeParse(content).success
    )
      return null;
    const sourceSheet = sheetOf(
      spreadsheetPersistenceJson(request.source),
      request.sheetId,
    );
    const original = memberOf(sourceSheet, request.sourceObjectId);
    const copied = memberOf(
      sheetOf(content, request.sheetId),
      receipt.copiedObjectId,
    );
    if (
      !sourceSheet ||
      !original ||
      !copied ||
      copied.recipe ||
      !boundCopy(copied.object, receipt) ||
      (original.recipe ? "chart" : original.object.type) !==
        receipt.objectKind ||
      (request.source.sheets as unknown[]).some((sheet) =>
        memberOf(record(sheet), receipt.copiedObjectId),
      ) ||
      !equal(
        withoutCopy(
          content,
          request.sheetId,
          receipt.copiedObjectId,
          sourceSheet,
        ),
        authoring(request.source),
      )
    )
      return null;
    const originalSource = record(original.object.source),
      copiedSource = record(copied.object.source)!;
    if (
      originalSource &&
      (originalSource.drawingPart !== receipt.drawingPart ||
        (receipt.objectKind === "shape"
          ? originalSource.nonVisualId !== receipt.sourceNonVisualId
          : receipt.partMap[String(originalSource.part)] !==
              copiedSource.part ||
            receipt.relationshipIdMap[String(originalSource.relationshipId)] !==
              copiedSource.relationshipId))
    )
      return null;
    const draft = spreadsheetPersistenceJson(request.draft ?? request.source);
    const reviewed = memberOf(
      sheetOf(draft, request.sheetId),
      request.sourceObjectId,
    );
    if (
      !reviewed ||
      reviewed.recipe !== original.recipe ||
      !equal(reviewed.object.source, original.object.source)
    )
      return null;
    const unchangedDraft = authoring(draft),
      draftSheet = sheetOf(unchangedDraft, request.sheetId)!;
    const key = reviewed.recipe ? "charts" : "objects";
    const members = draftSheet[key] as Json[];
    draftSheet[key] = members.map((member) =>
      member.id === request.sourceObjectId
        ? structuredClone(original.object)
        : member,
    );
    if (!equal(unchangedDraft, authoring(request.source))) return null;
    if (!original.recipe) {
      const expected =
        receipt.objectKind === "chart"
          ? reviewedChartMirrors(
              original.object,
              reviewed.object,
              copied.object.chartType,
            )
          : { ...reviewed.object };
      if (!expected) return null;
      const before = { ...original.object };
      delete expected.id;
      delete expected.source;
      delete before.id;
      delete before.source;
      if (!reviewedIntent(before, expected, copied.object)) return null;
    } else {
      // Range/type changes are outside this backend's source-preserving recipe
      // copy contract. Do not invent a second range/XY/combo converter here.
      if (
        original.object.type === "camera" ||
        original.object.pivotSource !== undefined ||
        ["type", "range", "sourceSheetId", "seriesInColumns", "noCategories", "noSeriesNames"].some(
          (key) => !equal(original.object[key], reviewed.object[key]),
        ) ||
        copied.object.chartType !== reviewed.object.type ||
        !Array.isArray(copied.object.seriesRefs) ||
        copied.object.seriesRefs.length === 0
      )
        return null;
      const excluded = new Set([
        "id",
        "type",
        "range",
        "sourceSheetId",
        "seriesInColumns",
        "noCategories",
        "noSeriesNames",
        "anchor",
        "width",
        "height",
        "x",
        "y",
      ]);
      const expectedRecipe = reviewedChartMirrors(
        original.object,
        reviewed.object,
        copied.object.chartType,
      );
      if (!expectedRecipe) return null;
      const wantsTitle =
        typeof reviewed.object.title === "string" &&
        reviewed.object.title.trim().length > 0;
      // A recipe's literal "column" is a real requested title, unlike the
      // native reader's identical fallback label after c:title was removed.
      if (copied.object.nativeTitlePresent !== wantsTitle) return null;
      if (
        Object.keys(original.object).some(
          (key) =>
            !excluded.has(key) &&
            !Object.hasOwn(reviewed.object, key) &&
            (key === "title"
              ? copied.object.nativeTitlePresent !== false
              : Object.hasOwn(copied.object, key)),
        )
      )
        return null;
      if (
        !Object.entries(expectedRecipe).every(
          ([key, item]) =>
            excluded.has(key) || subset(item, copied.object[key]),
        )
      )
        return null;
      const anchor = record(copied.object.anchor),
        intended = record(reviewed.object.anchor) ?? {};
      if (
        !anchor ||
        anchor.widthEmu !==
          Math.round(Number(reviewed.object.width ?? 320) * 9525) ||
        anchor.heightEmu !==
          Math.round(Number(reviewed.object.height ?? 200) * 9525)
      )
        return null;
      if (reviewed.object.x !== undefined || reviewed.object.y !== undefined) {
        if (
          anchor.kind !== "absolute" ||
          anchor.xEmu !== Math.round(Number(reviewed.object.x) * 9525) ||
          anchor.yEmu !== Math.round(Number(reviewed.object.y) * 9525)
        )
          return null;
      } else if (
        anchor.kind !== "oneCell" ||
        record(anchor.from)?.row !== (intended.row ?? 0) ||
        record(anchor.from)?.column !== (intended.col ?? 0)
      )
        return null;
    }
    return {
      content,
      receipt,
      insertedSheetIds: [],
      insertedObjectIds: [receipt.copiedObjectId],
    };
  } catch {
    return null;
  }
}

/** The host must flush and bind a current source/actor before requesting. */
export async function requestSpreadsheetNativeObjectCopy(
  request: SpreadsheetNativeObjectCopyRequest,
  signal?: AbortSignal,
): Promise<SpreadsheetNativeObjectCopyPlan> {
  let draft: Json | undefined;
  if (request.draft) {
    draft = structuredClone(spreadsheetPersistenceJson(request.draft));
    for (const key of controls) {
      if (!equal(draft[key], request.source[key]))
        throw new Error("native_xlsx_object_copy_draft_receipt_changed");
      delete draft[key];
    }
  }
  const value = await apiPost<unknown>(
    `/api/v1/pages/${encodeURIComponent(request.pageId)}/native-xlsx-object-copy-plan`,
    {
      expected_content_version: request.contentVersion,
      sheet_id: request.sheetId,
      source_object_id: request.sourceObjectId,
      ...(draft ? { draft_content: draft } : {}),
    },
    { signal },
  );
  const plan = readSpreadsheetNativeObjectCopyPlan(value, request);
  if (!plan) throw new Error("native_xlsx_object_copy_response_mismatch");
  return plan;
}

function savedParts(value: unknown) {
  const page = record(value),
    content = record(page?.content_json);
  const direct = record(page?.editable_page_patchback),
    nested = record(record(page?.ai_status_json)?.editable_page_patchback);
  if (direct && nested && !equal(direct, nested)) return null;
  const marker = direct ?? nested,
    copy = record(marker?.native_object_copy);
  return page &&
    content &&
    marker &&
    copy &&
    marker.native_worksheet_copy === undefined
    ? { page, content, marker, copy }
    : null;
}
function savedIdentity(
  parts: NonNullable<ReturnType<typeof savedParts>>,
  pageId: string,
  previousVersion: number,
  resourceId: string,
): boolean {
  const { page, marker, copy } = parts;
  return (
    page.id === pageId &&
    page.content_version === previousVersion + 1 &&
    marker.status === "succeeded" &&
    marker.adapter === "xlsx" &&
    marker.loss_entry_count === 0 &&
    marker.resource_id === resourceId &&
    identifier.safeParse(marker.resource_version_id).success &&
    digest.safeParse(marker.output_sha256).success &&
    copy.saved_page_id === pageId &&
    copy.saved_content_version === previousVersion + 1
  );
}
export function readSpreadsheetNativeObjectCopySavedReceipt(
  value: unknown,
  request: SpreadsheetNativeObjectCopyRequest,
  plan: SpreadsheetNativeObjectCopyPlan,
  actorIncarnation: string,
): SpreadsheetNativeCopySavedReceipt | null {
  try {
    const saved = savedParts(value);
    if (
      !saved ||
      !actorIncarnation ||
      plan.receipt.pageId !== request.pageId ||
      plan.receipt.sourceContentVersion !== request.contentVersion ||
      plan.receipt.sheetId !== request.sheetId ||
      plan.receipt.sourceObjectId !== request.sourceObjectId ||
      !equal(plan.insertedSheetIds, []) ||
      !equal(plan.insertedObjectIds, [plan.receipt.copiedObjectId]) ||
      !savedIdentity(
        saved,
        request.pageId,
        request.contentVersion,
        plan.receipt.sourceResourceId,
      ) ||
      saved.copy.inserted !== true ||
      saved.copy.history !== undefined ||
      !equal(saved.copy.receipt, plan.receipt) ||
      !equal(authoring(saved.content, false), authoring(plan.content, false))
    )
      return null;
    return {
      schemaVersion: "xlsx-native-copy-saved-v1",
      savedPageId: request.pageId,
      sourceContentVersion: request.contentVersion,
      savedVersion: request.contentVersion + 1,
      actorIncarnation,
      planId: plan.receipt.planId,
      insertedSheetIds: [],
      insertedObjectIds: [plan.receipt.copiedObjectId],
      sourcePreserved: true,
      native: {
        resourceId: plan.receipt.sourceResourceId,
        resourceVersionId: saved.marker.resource_version_id as string,
        sourcePackageSha256: plan.receipt.sourcePackageSha256,
        outputPackageSha256: saved.marker.output_sha256 as string,
        lossCount: 0,
      },
    };
  } catch {
    return null;
  }
}

export function createSpreadsheetNativeObjectCopyHistoryStep(
  plan: SpreadsheetNativeObjectCopyPlan,
  saved: SpreadsheetNativeCopySavedReceipt,
  actorId: string,
): SpreadsheetNativeObjectCopyHistoryStep | null {
  const object = memberOf(
    sheetOf(plan.content, plan.receipt.sheetId),
    plan.receipt.copiedObjectId,
  );
  if (
    !object ||
    object.recipe ||
    saved.planId !== plan.receipt.planId ||
    saved.savedPageId !== plan.receipt.pageId ||
    saved.sourceContentVersion !== plan.receipt.sourceContentVersion ||
    saved.savedVersion !== plan.receipt.sourceContentVersion + 1 ||
    saved.native.resourceId !== plan.receipt.sourceResourceId ||
    !saved.actorIncarnation ||
    !actorId ||
    !equal(saved.insertedSheetIds, []) ||
    !equal(saved.insertedObjectIds, [plan.receipt.copiedObjectId])
  )
    return null;
  return {
    receipt: structuredClone(plan.receipt),
    copiedObject: structuredClone(object.object),
    copyExportVersionId: saved.native.resourceVersionId,
    actorIncarnation: saved.actorIncarnation,
    actorId,
  };
}
function remappedObject(
  step: SpreadsheetNativeObjectCopyHistoryStep,
  history: z.infer<typeof historySchema>,
): Json | null {
  if (
    !equal(
      Object.keys(history.partMap).sort(),
      Object.keys(step.receipt.partMap).sort(),
    ) ||
    !equal(
      Object.keys(history.relationshipIdMap).sort(),
      Object.keys(step.receipt.relationshipIdMap).sort(),
    )
  )
    return null;
  const object = structuredClone(step.copiedObject),
    source = record(object.source);
  if (!source || !boundCopy(object, step.receipt)) return null;
  for (const [field, initial, next] of [
    ["part", step.receipt.partMap, history.partMap],
    [
      "relationshipId",
      step.receipt.relationshipIdMap,
      history.relationshipIdMap,
    ],
  ] as const) {
    if (source[field] === undefined) continue;
    const keys = Object.keys(initial).filter(
      (key) => initial[key] === source[field],
    );
    if (keys.length !== 1) return null;
    source[field] = next[keys[0]];
  }
  if (source.nonVisualId !== undefined)
    source.nonVisualId = history.copiedNonVisualId;
  return object;
}
export function readSpreadsheetNativeObjectCopyHistoryPlan(
  value: unknown,
  step: SpreadsheetNativeObjectCopyHistoryStep,
  operation: "undo" | "redo",
  currentVersion: number,
  source: Json,
): SpreadsheetNativeObjectCopyHistoryPlan | null {
  try {
    const response = record(value),
      content = record(response?.content);
    const parsed = historySchema.safeParse(response?.receipt),
      insertion = receiptSchema.safeParse(step.receipt);
    if (
      !content ||
      !parsed.success ||
      !insertion.success ||
      !version.safeParse(currentVersion).success ||
      !spreadsheetPageContentSchema.safeParse(content).success
    )
      return null;
    const receipt = parsed.data,
      origin = insertion.data;
    if (
      receipt.operation !== operation ||
      receipt.pageId !== origin.pageId ||
      receipt.sourceContentVersion !== currentVersion ||
      receipt.copyExportVersionId !== step.copyExportVersionId ||
      receipt.copyPlanId !== origin.planId ||
      receipt.copiedNonVisualId === origin.sourceNonVisualId ||
      !equal(content.nativeObjectCopyHistory, receipt) ||
      controls.some(
        (key) =>
          key !== "nativeObjectCopyHistory" && Object.hasOwn(content, key),
      )
    )
      return null;
    const oldSheet = sheetOf(source, origin.sheetId),
      nextSheet = sheetOf(content, origin.sheetId);
    const before = memberOf(oldSheet, origin.copiedObjectId),
      after = memberOf(nextSheet, origin.copiedObjectId);
    const expected = remappedObject(step, receipt);
    if (!oldSheet || !nextSheet || !expected) return null;
    if (operation === "undo") {
      if (
        !before ||
        before.recipe ||
        after ||
        !equal(before.object, expected) ||
        !equal(
          withoutCopy(source, origin.sheetId, origin.copiedObjectId, nextSheet),
          authoring(content),
        )
      )
        return null;
    } else if (
      before ||
      !after ||
      after.recipe ||
      !equal(after.object, expected) ||
      !equal(
        withoutCopy(content, origin.sheetId, origin.copiedObjectId, oldSheet),
        authoring(source),
      )
    )
      return null;
    return { content, receipt };
  } catch {
    return null;
  }
}
export async function requestSpreadsheetNativeObjectCopyHistory(
  step: SpreadsheetNativeObjectCopyHistoryStep,
  operation: "undo" | "redo",
  currentVersion: number,
  source: Json,
  signal?: AbortSignal,
): Promise<SpreadsheetNativeObjectCopyHistoryPlan> {
  const response = await apiPost<unknown>(
    `/api/v1/pages/${encodeURIComponent(step.receipt.pageId)}/native-xlsx-object-copy-history-plan`,
    {
      expected_content_version: currentVersion,
      copy_export_version_id: step.copyExportVersionId,
      operation,
    },
    { signal },
  );
  const plan = readSpreadsheetNativeObjectCopyHistoryPlan(
    response,
    step,
    operation,
    currentVersion,
    source,
  );
  if (!plan)
    throw new Error("native_xlsx_object_copy_history_response_mismatch");
  return plan;
}
export function spreadsheetNativeObjectCopyHistorySaved(
  value: unknown,
  step: SpreadsheetNativeObjectCopyHistoryStep,
  currentVersion: number,
  plan: SpreadsheetNativeObjectCopyHistoryPlan,
  actorIncarnation: string,
): boolean {
  try {
    const saved = savedParts(value);
    return Boolean(
      saved &&
      actorIncarnation === step.actorIncarnation &&
      actorIncarnation &&
      savedIdentity(
        saved,
        step.receipt.pageId,
        currentVersion,
        step.receipt.sourceResourceId,
      ) &&
      saved.copy.inserted === false &&
      equal(saved.copy.receipt, step.receipt) &&
      equal(saved.copy.history, plan.receipt) &&
      saved.marker.output_sha256 === plan.receipt.outputPackageSha256 &&
      equal(authoring(saved.content, false), authoring(plan.content, false)),
    );
  } catch {
    return false;
  }
}
