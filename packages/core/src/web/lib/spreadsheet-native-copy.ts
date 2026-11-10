/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { NativeDocumentCommand } from "../components/console/editor/native-document/documentCommands";
import type {
  CopilotSpreadsheetContent,
  CopilotSpreadsheetDelta,
  CopilotSpreadsheetSelection,
} from "../__cut__/mfc2dd2a6e6";

export type SpreadsheetNativeCopyPlacement =
  | { kind: "worksheet"; afterSheetId: string }
  | { kind: "object"; sheetId: string; targetKind: "object" | "chart"; targetId: string };

/** Native plans describe their resulting focus, but a surviving member view
 * must stay put while that plan is saved. Changing it early retires the hot
 * shard generation. The host applies the planned focus after acknowledgement.
 */
export function withCurrentSpreadsheetCopyView(
  planned: Record<string, unknown>,
  current: Record<string, unknown>,
): Record<string, unknown> {
  const content = { ...planned };
  const active = current.activeSheetId;
  if (typeof active !== "string") delete content.activeSheetId;
  else if (Array.isArray(planned.sheets) && planned.sheets.some(sheet =>
    sheet !== null && typeof sheet === "object" && sheet.id === active)) {
    content.activeSheetId = active;
  }
  // If Undo removes the active sheet, retain the server's surviving fallback.
  return content;
}

/** Host-only contract: callbacks never enter the API/model payload. The host
 * authenticates the saved source and materializes the reviewed draft into an
 * independent native copy before returning a successful save receipt.
 */
export interface SpreadsheetNativeCopyInput {
  pageId: string;
  command: NativeDocumentCommand;
  selection: CopilotSpreadsheetSelection;
  source: CopilotSpreadsheetContent;
  draft: CopilotSpreadsheetContent;
  delta: CopilotSpreadsheetDelta;
  actorIncarnation: string;
  placement: SpreadsheetNativeCopyPlacement;
  isCurrent(): boolean;
  isActorCurrent(): boolean;
}

export interface SpreadsheetNativeCopySavedReceipt {
  schemaVersion: "xlsx-native-copy-saved-v1";
  savedPageId: string;
  sourceContentVersion: number;
  savedVersion: number;
  actorIncarnation: string;
  planId: string;
  insertedSheetIds: string[];
  insertedObjectIds: string[];
  sourcePreserved: true;
  native: {
    resourceId: string;
    resourceVersionId: string;
    sourcePackageSha256: string;
    outputPackageSha256: string;
    lossCount: 0;
  };
}

export type SpreadsheetNativeCopyResult =
  | { ok: true; receipt: SpreadsheetNativeCopySavedReceipt }
  | { ok: false; message?: string; plainCopied?: true };

export type ApplySpreadsheetNativeCopy = (
  input: SpreadsheetNativeCopyInput,
) => Promise<SpreadsheetNativeCopyResult>;
