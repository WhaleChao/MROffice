/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  mergePageDocumentDelta,
  type PageContentDocument,
} from "../hooks/usePageContentSync";
import { getPage } from "./notebook-sdk";
import { isApiRequestError } from "./api";
import { pageDocumentJsonEquals as equal } from "./page-save-acknowledgement";
import { spreadsheetPersistenceJson } from "./spreadsheet/persistence-json";
import {
  compactHydratedContentBlankCellFormats,
  hydrateContentBlankCellFormats,
} from "./spreadsheet/blank-cell-formats";
import {
  spreadsheetCopilotCreatedChartIds,
  spreadsheetCopilotWorkbookDelta,
  type CopilotSpreadsheetContent,
  type NativeCopilotCellValuePlanner,
} from "../__cut__/mfc2dd2a6e6";
import type {
  SpreadsheetNativeCopyInput,
  SpreadsheetNativeCopyResult,
  SpreadsheetNativeCopySavedReceipt,
} from "./spreadsheet-native-copy";
import { withCurrentSpreadsheetCopyView } from "./spreadsheet-native-copy";
import type { PageSaveWithAcknowledgementResult } from "../hooks/useAutosavedPageDocument";
import { copilotAuthoredWorkbook } from "../__cut__/m32631ef371";
import {
  requestSpreadsheetNativeWorksheetCopy,
  readSpreadsheetNativeWorksheetCopySavedReceipt,
  type SpreadsheetNativeWorksheetCopyPlan,
} from "./spreadsheet-native-worksheet-copy";

export interface SpreadsheetNativeCopyHost {
  /** The page's `console-notebooks` translator for the refusal lines. */
  translate: (key: string) => string;
  isWritable(): boolean;
  prepareSource?(
    page: Awaited<ReturnType<typeof getPage>>,
    isCurrent: () => boolean,
  ): Promise<boolean>;
  readSource(): CopilotSpreadsheetContent;
  normalizeSavedContent(
    content: Record<string, unknown>,
  ): CopilotSpreadsheetContent;
  copiedName: string;
  planValues: NativeCopilotCellValuePlanner;
  save(content: Record<string, unknown>): Promise<unknown>;
  saveWithAcknowledgement?(
    content: Record<string, unknown>,
  ): Promise<PageSaveWithAcknowledgementResult>;
  accept(
    plan: SpreadsheetNativeWorksheetCopyPlan,
    saved: unknown,
    receipt: SpreadsheetNativeCopySavedReceipt,
  ): void;
}

const persisted = (content: CopilotSpreadsheetContent) =>
  spreadsheetPersistenceJson(content);

/** Replay the already reviewed intent onto the raw saved projection. Renderer
 * defaults are not edits and must not become new native package authority.
 * The editor's source and draft hold formatted blanks as ordinary cells, so
 * the replay happens on the saved page in that same shape and the result is
 * stored as ranges again, untouched ranges exactly as saved.
 */
export function projectSpreadsheetCopyDraftToSavedSource(
  saved: Record<string, unknown>,
  source: CopilotSpreadsheetContent,
  draft: CopilotSpreadsheetContent,
): Record<string, unknown> {
  const hydrated = hydrateContentBlankCellFormats(spreadsheetPersistenceJson(saved));
  return compactHydratedContentBlankCellFormats(
    mergePageDocumentDelta(
      hydrated as PageContentDocument,
      persisted(draft) as PageContentDocument,
      persisted(source) as PageContentDocument,
    ),
    hydrated,
  );
}






function reviewedIntentOnCopiedSheet(
  input: SpreadsheetNativeCopyInput,
): SpreadsheetNativeCopyInput["delta"] | null {
  const intent = input.delta.intent;
  if (!intent) return null;
  const ids = (content: CopilotSpreadsheetContent) => content.sheets.map((sheet) => sheet.id);
  const others = (content: CopilotSpreadsheetContent) =>
    copilotAuthoredWorkbook({
      ...content,
      activeSheetId: input.source.activeSheetId,
      sheets: content.sheets.filter((sheet) => sheet.id !== input.selection.sheet_id),
    });
  return equal(ids(input.draft), ids(input.source)) &&
    equal(copilotAuthoredWorkbook(intent.expected), copilotAuthoredWorkbook(input.draft)) &&
    equal(others(input.draft), others(input.source))
    ? input.delta
    : null;
}

export async function applySpreadsheetNativeWorksheetCopy(
  input: SpreadsheetNativeCopyInput,
  host: SpreadsheetNativeCopyHost,
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
  if (input.placement.kind !== "worksheet") {
    return fail("spreadsheet.nativeCopy.objectNotSupported");
  }
  if (input.placement.afterSheetId !== input.selection.sheet_id || !current()) {
    return fail("spreadsheet.nativeCopy.sheetChanged");
  }
  const delta = input.delta.intent
    ? reviewedIntentOnCopiedSheet(input)
    : spreadsheetCopilotWorkbookDelta(
        input.selection,
        input.source,
        input.draft,
        host.planValues,
        spreadsheetCopilotCreatedChartIds(input.command, input.selection),
      );
  if (!delta || !equal(delta, input.delta)) {
    return fail("spreadsheet.nativeCopy.changesNotVerified");
  }
  try {
    const sourcePage = await getPage(input.pageId);
    if (
      !current() ||
      sourcePage.id !== input.pageId ||
      sourcePage.page_type !== "spreadsheet" ||
      sourcePage.content_version !== input.selection.version ||
      !sourcePage.content_json ||
      !equal(
        persisted(host.normalizeSavedContent(sourcePage.content_json)),
        persisted(input.source),
      )
    ) {
      return fail("spreadsheet.nativeCopy.sheetNotSynchronized");
    }
    if (host.prepareSource && !(await host.prepareSource(sourcePage, current)))
      return fail("spreadsheet.nativeCopy.sourceUnavailable");
    const request = {
      pageId: input.pageId,
      contentVersion: input.selection.version,
      sourceSheetId: input.selection.sheet_id,
      copiedName: host.copiedName,
      source: sourcePage.content_json,
      draft: projectSpreadsheetCopyDraftToSavedSource(
        sourcePage.content_json,
        input.source,
        input.draft,
      ),
    };
    const plan = await requestSpreadsheetNativeWorksheetCopy(request);
    if (!current()) return fail("spreadsheet.nativeCopy.sheetChangedNoCopy");
    // This is the existing normal saver, including Workbook V2's structural
    // transaction. The private preview is never written over the source first.
    // Keep the current member view while its new sibling is being saved.
    // Changing the hot sheet here retires the generation performing that save.
    // The host moves focus to the copy after the durable acknowledgement.
    const saveContent = withCurrentSpreadsheetCopyView(plan.content, input.source);
    const ownedSave = host.saveWithAcknowledgement
      ? await host.saveWithAcknowledgement(saveContent)
      : null;
    const saved = host.saveWithAcknowledgement
      ? ownedSave?.savedPage
      : await host.save(saveContent);
    const canonicalAcknowledgement = ownedSave?.acknowledgement;
    if (!input.isActorCurrent()) return { ok: false };
    let acknowledged = saved;
    let receipt = readSpreadsheetNativeWorksheetCopySavedReceipt(
      acknowledged,
      request,
      plan,
      input.actorIncarnation,
      canonicalAcknowledgement,
    );
    if (!receipt) {
      // V2 may invalidate its local projection while a committed structural
      // save reopens the shard generation, returning null or a display carry.
      // Reconcile that inconclusive result against the exact durable native
      // receipt; never resend the plan or accept a generic successful GET.
      acknowledged = await getPage(input.pageId);
      if (!input.isActorCurrent()) return { ok: false };
      receipt = readSpreadsheetNativeWorksheetCopySavedReceipt(
        acknowledged,
        request,
        plan,
        input.actorIncarnation,
        canonicalAcknowledgement,
      );
    }
    if (!receipt) return fail("spreadsheet.nativeCopy.sheetCopyNotConfirmed");
    const canonicalContent = (
      acknowledged as { content_json: Record<string, unknown> }
    ).content_json;
    host.accept({ ...plan, content: canonicalContent }, acknowledged, receipt);
    return { ok: true, receipt };
  } catch (error) {
    if (
      isApiRequestError(error) &&
      error.code === "native_xlsx_copy_projection_limit"
    ) {
      return fail("spreadsheet.nativeCopy.imageCapacity");
    }
    try {
      // The member sees "not saved"; the step that threw is for the console.
      console.error(
        `[sheet-native-copy] not saved: ${
          error instanceof Error ? `${error.name}: ${error.message}` : String(error)
        }`,
      );
    } catch {
      /* a diagnostic never costs the refusal */
    }
    return fail("spreadsheet.nativeCopy.sheetCopyNotSaved");
  }
}
