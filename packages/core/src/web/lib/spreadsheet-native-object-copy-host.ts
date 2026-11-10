/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { getPage } from "./notebook-sdk";
import { getWorkbookCollabV2 } from "./page-collab-sdk";
import { pageDocumentJsonEquals as equal } from "./page-save-acknowledgement";
import { spreadsheetPersistenceJson } from "./spreadsheet/persistence-json";
import { spreadsheetPageContentSchema } from "./native-spreadsheet-schemas";
import { projectSpreadsheetCopyDraftToSavedSource } from "./spreadsheet-native-copy-host";
import {
  withCurrentSpreadsheetCopyView,
  type SpreadsheetNativeCopyInput,
  type SpreadsheetNativeCopyResult,
  type SpreadsheetNativeCopySavedReceipt,
} from "./spreadsheet-native-copy";
import {
  spreadsheetCopilotWorkbookDelta,
  type CopilotSpreadsheetContent,
  type NativeCopilotCellValuePlanner,
} from "../__cut__/mfc2dd2a6e6";
import {
  captureSavedPageHotCacheScope,
  isSavedPageHotCacheScopeCurrent,
} from "./saved-page-hot-cache";
import {
  requestSpreadsheetNativeObjectCopy,
  readSpreadsheetNativeObjectCopySavedReceipt,
  requestSpreadsheetNativeObjectCopyHistory,
  spreadsheetNativeObjectCopyHistorySaved,
  type SpreadsheetNativeObjectCopyPlan,
  type SpreadsheetNativeObjectCopyHistoryStep,
} from "./spreadsheet-native-object-copy";

type Json = Record<string, unknown>;
export function spreadsheetNativeObjectCopyAvailable<
  T extends {
    id: string;
    objects?: { id: string }[];
    charts?: { id: string }[];
  },
>(
  selection: SpreadsheetNativeCopyInput["selection"],
  sheets: T[],
  hasSource: boolean,
  isProtected: (sheet: T) => boolean,
  objectLimit: number,
): boolean {
  const target = selection.targets.length === 1 ? selection.targets[0] : null;
  const candidates = sheets.filter((sheet) => sheet.id === selection.sheet_id);
  if (
    !hasSource ||
    !target ||
    target.kind === "cell" ||
    candidates.length !== 1
  )
    return false;
  const sheet = candidates[0],
    members = target.kind === "chart" ? sheet.charts : sheet.objects;
  return (
    !isProtected(sheet) &&
    (sheet.objects?.length ?? 0) < objectLimit &&
    members?.filter((object) => object.id === target.object_id).length === 1
  );
}
export function reconcileSpreadsheetObjectCopyView<
  T extends { id: string; objects?: { id: string }[] },
>(
  live: T[],
  planned: T[],
  sheetId: string,
  objectId: string,
  operation: "insert" | "undo" | "redo",
): T[] | null {
  const indexes = live.flatMap((sheet, index) =>
    sheet.id === sheetId ? [index] : [],
  );
  if (indexes.length !== 1) return null;
  const index = indexes[0],
    sheet = live[index],
    objects = sheet.objects ?? [];
  const matches = objects.filter((object) => object.id === objectId);
  if (matches.length > 1) return null;
  let next = objects;
  if (operation === "undo") {
    if (!matches.length) return live;
    next = objects.filter((object) => object.id !== objectId);
  } else {
    if (matches.length) return live;
    const plans = planned.filter((candidate) => candidate.id === sheetId);
    const expected = plans.length === 1 ? (plans[0].objects ?? []) : [];
    const inserted = expected.filter((object) => object.id === objectId);
    if (inserted.length !== 1) return null;
    const position = expected.indexOf(inserted[0]);
    const predecessor = expected
      .slice(0, position)
      .reverse()
      .find((object) => objects.some((current) => current.id === object.id));
    const after = predecessor
      ? objects.findIndex((object) => object.id === predecessor.id) + 1
      : 0;
    next = [...objects.slice(0, after), inserted[0], ...objects.slice(after)];
  }
  return live.map((candidate, at) =>
    at === index ? { ...candidate, objects: next } : candidate,
  );
}
const record = (value: unknown): Json | null =>
  value !== null && typeof value === "object" && !Array.isArray(value)
    ? (value as Json)
    : null;
const persisted = (value: CopilotSpreadsheetContent) =>
  spreadsheetPersistenceJson(value);

export interface SpreadsheetNativeObjectCopyHost {
  /** The page's `console-notebooks` translator for the refusal lines. */
  translate: (key: string) => string;
  isWritable(): boolean;
  readSource(): CopilotSpreadsheetContent;
  normalizeSavedContent(content: Json): CopilotSpreadsheetContent;
  prepareSource?(
    page: Awaited<ReturnType<typeof getPage>>,
    isCurrent: () => boolean,
  ): Promise<boolean>;
  planValues: NativeCopilotCellValuePlanner;
  save(content: Json): Promise<unknown>;
  accept(
    plan: SpreadsheetNativeObjectCopyPlan,
    saved: unknown,
    receipt: SpreadsheetNativeCopySavedReceipt,
  ): boolean;
}

/** Office-like small nearby displacement, using only actual native markers.
 * Both two-cell endpoints move equally; dimensions, source parts and data
 * references stay untouched. The normal native saver must verify this edit.
 */
export function positionSpreadsheetObjectCopyNearby(
  plan: SpreadsheetNativeObjectCopyPlan,
): SpreadsheetNativeObjectCopyPlan | null {
  const content = structuredClone(plan.content);
  const sheets = Array.isArray(content.sheets)
    ? content.sheets
        .map(record)
        .filter((sheet) => sheet?.id === plan.receipt.sheetId)
    : [];
  const sheet = sheets.length === 1 ? sheets[0] : null;
  const objects = Array.isArray(sheet?.objects)
    ? sheet.objects
        .map(record)
        .filter((object) => object?.id === plan.receipt.copiedObjectId)
    : [];
  const object = objects.length === 1 ? objects[0] : null;
  const anchor = record(object?.anchor);
  if (!object || !anchor) return null;
  const move = (value: unknown, optional = false): number | null => {
    const previous = optional && value === undefined ? 0 : value;
    return typeof previous === "number" &&
      Number.isSafeInteger(previous) &&
      Number.isSafeInteger(previous + 190500)
      ? previous + 190500
      : null;
  };
  if (anchor.kind === "absolute") {
    const x = move(anchor.xEmu),
      y = move(anchor.yEmu);
    if (x === null || y === null) return null;
    anchor.xEmu = x;
    anchor.yEmu = y;
  } else if (anchor.kind === "oneCell" || anchor.kind === "twoCell") {
    for (const key of anchor.kind === "twoCell" ? ["from", "to"] : ["from"]) {
      const point = record(anchor[key]);
      if (
        !point ||
        !Number.isSafeInteger(point.row) ||
        !Number.isSafeInteger(point.column)
      )
        return null;
      const x = move(point.columnOffset, true),
        y = move(point.rowOffset, true);
      if (x === null || y === null) return null;
      point.columnOffset = x;
      point.rowOffset = y;
    }
  } else return null;
  return spreadsheetPageContentSchema.safeParse(content).success
    ? { ...plan, content }
    : null;
}

export async function applySpreadsheetNativeObjectCopy(
  input: SpreadsheetNativeCopyInput,
  host: SpreadsheetNativeObjectCopyHost,
): Promise<SpreadsheetNativeCopyResult> {
  const fail = (key: string): SpreadsheetNativeCopyResult => ({
    ok: false,
    message: host.translate(key),
  });
  const current = () =>
    input.isCurrent() &&
    input.isActorCurrent() &&
    host.isWritable() &&
    equal(persisted(host.readSource()), persisted(input.source));
  const target =
    input.selection.targets.length === 1 ? input.selection.targets[0] : null;
  if (
    input.placement.kind !== "object" ||
    !target ||
    target.kind === "cell" ||
    input.placement.sheetId !== input.selection.sheet_id ||
    input.placement.targetKind !== target.kind ||
    input.placement.targetId !== target.object_id ||
    !current()
  )
    return fail("spreadsheet.nativeCopy.objectChanged");
  const delta = spreadsheetCopilotWorkbookDelta(
    input.selection,
    input.source,
    input.draft,
    host.planValues,
  );
  if (
    !delta ||
    !equal(delta, input.delta) ||
    delta.cells.length ||
    delta.objects.length !== 1 ||
    delta.objects[0].id !== target.object_id ||
    !delta.objects[0].after ||
    delta.createdCharts?.length
  )
    return fail("spreadsheet.nativeCopy.objectEditNotVerified");
  try {
    const page = await getPage(input.pageId);
    if (
      !current() ||
      page.id !== input.pageId ||
      page.page_type !== "spreadsheet" ||
      page.content_version !== input.selection.version ||
      !page.content_json ||
      !equal(
        persisted(host.normalizeSavedContent(page.content_json)),
        persisted(input.source),
      )
    )
      return fail("spreadsheet.nativeCopy.objectNotSynchronized");
    if (host.prepareSource && !(await host.prepareSource(page, current)))
      return fail("spreadsheet.nativeCopy.sourceUnavailable");
    const actor = await getWorkbookCollabV2(input.pageId);
    if (
      !current() ||
      actor.actorMembershipIncarnation !== input.actorIncarnation
    )
      return { ok: false };
    const request = {
      pageId: input.pageId,
      contentVersion: input.selection.version,
      sheetId: input.selection.sheet_id,
      sourceObjectId: target.object_id,
      source: page.content_json,
      draft: projectSpreadsheetCopyDraftToSavedSource(
        page.content_json,
        input.source,
        input.draft,
      ),
    };
    const prepared = await requestSpreadsheetNativeObjectCopy(request);
    if (!current()) return fail("spreadsheet.nativeCopy.objectChangedNoCopy");
    const plan = positionSpreadsheetObjectCopyNearby(prepared);
    if (!plan) return fail("spreadsheet.nativeCopy.copyPositionNotVerified");
    let saved = await host.save(
      withCurrentSpreadsheetCopyView(plan.content, input.source),
    );
    if (!input.isActorCurrent()) return { ok: false };
    let receipt = readSpreadsheetNativeObjectCopySavedReceipt(
      saved,
      request,
      plan,
      input.actorIncarnation,
    );
    if (!receipt) {
      // Inconclusive V2 acknowledgement is reconciled by reading once, never
      // by resending a structural mutation that may already be committed.
      saved = await getPage(input.pageId);
      if (!input.isActorCurrent()) return { ok: false };
      receipt = readSpreadsheetNativeObjectCopySavedReceipt(
        saved,
        request,
        plan,
        input.actorIncarnation,
      );
    }
    if (!receipt || !host.accept(plan, saved, receipt))
      return fail("spreadsheet.nativeCopy.objectCopyNotConfirmed");
    return { ok: true, receipt };
  } catch {
    return fail("spreadsheet.nativeCopy.objectCopyNotSaved");
  }
}

export interface SpreadsheetNativeObjectHistoryHost {
  isCurrent(): boolean;
  readSource(): CopilotSpreadsheetContent;
  normalizeSavedContent(content: Json): CopilotSpreadsheetContent;
  flush(): Promise<unknown>;
  save(content: Json): Promise<unknown>;
  accept(content: Json, savedVersion: number): boolean;
}
export async function applySpreadsheetNativeObjectCopyHistory(
  step: SpreadsheetNativeObjectCopyHistoryStep,
  operation: "undo" | "redo",
  host: SpreadsheetNativeObjectHistoryHost,
): Promise<boolean> {
  const actorScope = captureSavedPageHotCacheScope();
  const current = () =>
    host.isCurrent() && isSavedPageHotCacheScopeCurrent(actorScope);
  if (!current() || !(await host.flush()) || !current()) return false;
  const source = persisted(host.readSource());
  const unchanged = () =>
    current() && equal(persisted(host.readSource()), source);
  const page = await getPage(step.receipt.pageId);
  if (
    !page.content_json ||
    !Number.isSafeInteger(page.content_version) ||
    !unchanged() ||
    !equal(persisted(host.normalizeSavedContent(page.content_json)), source)
  )
    return false;
  const actor = await getWorkbookCollabV2(step.receipt.pageId);
  if (
    !unchanged() ||
    actor.actorMembershipIncarnation !== step.actorIncarnation
  )
    return false;
  const version = page.content_version!;
  const plan = await requestSpreadsheetNativeObjectCopyHistory(
    step,
    operation,
    version,
    page.content_json,
  );
  if (!unchanged()) return false;
  const saved = await host.save(
    withCurrentSpreadsheetCopyView(plan.content, source),
  );
  if (!current()) return false;
  if (
    !spreadsheetNativeObjectCopyHistorySaved(
      saved,
      step,
      version,
      plan,
      step.actorIncarnation,
    )
  ) {
    const durable = await getPage(step.receipt.pageId);
    if (
      !current() ||
      !spreadsheetNativeObjectCopyHistorySaved(
        durable,
        step,
        version,
        plan,
        step.actorIncarnation,
      )
    )
      return false;
  }
  return host.accept(plan.content, version + 1);
}
