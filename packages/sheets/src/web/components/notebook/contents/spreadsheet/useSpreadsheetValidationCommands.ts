"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { validationDefinePlan, validationsAfterDefine, validationsAfterClear, type ValidationDraft } from "../../../../lib/spreadsheet/validation-buttons";
import { useCallback, type RefObject } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { type SpreadsheetDataValidationPatchbackBlocker, spreadsheetDataValidationWriterReceiptMatches, spreadsheetDataValidationPatchbackRuleBlocker, spreadsheetDataValidationPatchbackSelectionSupported } from "../../../../lib/spreadsheet/visible-command-authority";
import { shiftFormula } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import { type ValidationRange, normalizeRules, rangeHasRule, rebaseValidationRule, sameValidationSettingsAs, type ValidationRule, withUntouchedRulesAsStored } from "@mrnote/core/web/lib/spreadsheet/data-validation.ts";
import { type rangeHasLockedCell as rangeHasLockedCellWithProtection } from "@mrnote/core/web/lib/spreadsheet/protection.ts";
import { type SheetModel } from "./spreadsheetSheetModel";
import { type DataValidationDialogAuthority, type PageSetupSelectionSnapshot } from "./spreadsheetMutationReceipts";
import { type SheetViewMappedCellMutation, type SheetViewMappedDrawingMutation, type SheetViewMappedLineMutation, type SheetViewMappedPageMutation } from "./spreadsheetSheetViews";
import { type Dispatch, type SetStateAction } from "react";

/** What useSpreadsheetValidationCommands reads from SpreadsheetWindow's render scope. */
export interface SpreadsheetValidationCommandsInputs {
  buildValidationDraftRule: (range: ValidationRange) => ValidationRule | null;
  dataValidationMutationCommandRef: RefObject<DataValidationDialogAuthority | null>;
  isSheetProtectionEffectiveRef: RefObject<(candidate: SheetModel | null | undefined) => boolean>;
  liveActiveSheetIndexRef: RefObject<number>;
  liveSheetsRef: RefObject<SheetModel[]>;
  rangeHasLockedCellRef: RefObject<(candidate: SheetModel, range: Parameters<typeof rangeHasLockedCellWithProtection>[1]) => boolean>;
  readPageSetupSelectionSnapshot: () => PageSetupSelectionSnapshot | null;
  replaceActiveSheet: (nextSheet: SheetModel, nextTitle?: string, calculationMutation?: "content" | "coordinate", sheetViewMutation?: SheetViewMappedCellMutation | SheetViewMappedDrawingMutation | SheetViewMappedLineMutation | SheetViewMappedPageMutation, followOnOtherSheets?: (candidate: SheetModel) => SheetModel) => SheetModel[];
  setValidationError: Dispatch<SetStateAction<string | null>>;
  spreadsheetNativeWriteRoute: "patchback" | "from-scratch";
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  valApplyToSame: boolean;
  validationDraftFields: ValidationDraft;
  validationOpenedRule: ValidationRule | null;
}

/** Moved verbatim from SpreadsheetWindow (called at the same position, so hook order is unchanged). */
export function useSpreadsheetValidationCommands({
  buildValidationDraftRule,
  dataValidationMutationCommandRef,
  isSheetProtectionEffectiveRef,
  liveActiveSheetIndexRef,
  liveSheetsRef,
  rangeHasLockedCellRef,
  readPageSetupSelectionSnapshot,
  replaceActiveSheet,
  setValidationError,
  spreadsheetNativeWriteRoute,
  t,
  valApplyToSame,
  validationDraftFields,
  validationOpenedRule,
}: SpreadsheetValidationCommandsInputs) {
  // Why a rule cannot be written into the imported workbook: an MRAI-only
  // rule family, or Excel's own 255-character formula limit.
  const validationPatchbackBlockerMessage = useCallback(
    (blocker: SpreadsheetDataValidationPatchbackBlocker): string => {
      switch (blocker) {
        case "dependent-list":
          return t("spreadsheet.uiCopy.dependentDropdownsHaveNoExcelEquivalent");
        case "date-limits":
          return t("spreadsheet.uiCopy.excelDateRulesDontStoreWeekdayOrTodayOptions");
        case "formula-length":
          return t("spreadsheet.uiCopy.excelLimitsValidationFormulasTo255Characters");
        default:
          return t("spreadsheet.validationRuleInvalid");
      }
    },
    [t],
  );

  // Data validation (Excel/WPS parity). Define a rule on the exact selection
  // captured by the dialog authority. Apply re-resolves the workbook, sheet,
  // protection and selection epoch before entering the single history write.
  const defineValidationForSelection = useCallback((): boolean => {
    const authority = dataValidationMutationCommandRef.current;
    const liveSelection = readPageSetupSelectionSnapshot();
    const liveWorkbook = liveSheetsRef.current;
    const liveIndex = liveActiveSheetIndexRef.current;
    const liveSheet = liveWorkbook[liveIndex];
    if (
      !authority ||
      !liveSelection ||
      !liveSheet ||
      !spreadsheetDataValidationWriterReceiptMatches({
        expectedControlId: "XLSCTL.DATA.VALIDATION.MRAI.APPLY",
        receiptControlId: authority.controlId,
        sourceWorkbook: authority.workbook,
        sourceActiveSheetIndex: authority.activeSheetIndex,
        sourceSheet: authority.sheet,
        sourceSheetId: authority.sheetId,
        sourceSelectionEpoch: authority.selection.epoch,
        sourceRange: authority.selection.range,
        liveWorkbook,
        liveActiveSheetIndex: liveIndex,
        liveSheet,
        liveSheetId: liveSheet.id,
        liveSelectionEpoch: liveSelection.epoch,
        liveRange: liveSelection.range,
      }) ||
      isSheetProtectionEffectiveRef.current(liveSheet) ||
      rangeHasLockedCellRef.current(liveSheet, liveSelection.range)
    ) {
      setValidationError(t("spreadsheet.protectionBlocked"));
      return false;
    }
    const range: ValidationRange = { ...authority.selection.range };
    // Any value with no input message is Excel's "no validation": OK removes
    // the rule rather than writing an empty one (validation-buttons).
    const plan = validationDefinePlan(range, validationDraftFields);
    if (!plan) {
      setValidationError(t("spreadsheet.validationRuleInvalid"));
      return false;
    }
    const { rule } = plan;
    const liveValidations = normalizeRules(liveSheet.validations);
    if (spreadsheetNativeWriteRoute === "patchback") {
      const blocker = rule ? spreadsheetDataValidationPatchbackRuleBlocker(rule) : null;
      if (blocker) {
        setValidationError(validationPatchbackBlockerMessage(blocker));
        return false;
      }
      if (
        !spreadsheetDataValidationPatchbackSelectionSupported({
          selectionRange: range,
          rules: liveValidations,
        })
      ) {
        setValidationError(
          t("spreadsheet.uiCopy.overlappingValidationRulesStayAsTheFileHasThem"),
        );
        return false;
      }
    }
    // "Apply these changes to all other cells with the same settings": every
    // rule carrying the opened cell's settings takes the new ones on its own
    // cells, and the selection joins them.
    const sameAs = valApplyToSame ? validationOpenedRule : null;
    // A new rule on a range replaces the existing rule(s) on those cells only
    // (Excel applies one rule per cell; A1:A3 keeps A1 and A3 when A2 gets a
    // new rule), then prepends so it wins first-match lookup.
    // The draft's relative references read from the selection's top-left
    // (the dialog loaded them so): a same-settings rule elsewhere takes them
    // from its own top-left (B513).
    const buildFromSelection = (candidate: ValidationRange): ValidationRule | null => {
      const built = buildValidationDraftRule(candidate);
      if (!built || (candidate.r1 === range.r1 && candidate.c1 === range.c1)) return built;
      const atSelection = { ...built, range };
      const normalized = normalizeRules([atSelection])[0];
      return normalized ? rebaseValidationRule(normalized, candidate, shiftFormula) : built;
    };
    const next = validationsAfterDefine(
      liveValidations,
      range,
      plan,
      buildFromSelection,
      sameAs,
      shiftFormula,
    );
    if (next === "unchanged") {
      setValidationError(null);
      return true;
    }
    if (!next) {
      setValidationError(t("spreadsheet.validationRuleInvalid"));
      return false;
    }
    const nextSheet: SheetModel = { ...liveSheet };
    if (next.length > 0) nextSheet.validations = withUntouchedRulesAsStored(liveSheet.validations, next) as typeof nextSheet.validations;
    else delete nextSheet.validations;
    replaceActiveSheet(nextSheet);
    setValidationError(null);
    return true;
  }, [
    buildValidationDraftRule,
    readPageSetupSelectionSnapshot,
    replaceActiveSheet,
    spreadsheetNativeWriteRoute,
    t,
    valApplyToSame,
    validationDraftFields,
    validationOpenedRule,
    validationPatchbackBlockerMessage,
  ]);

  // Take the current selection out of every validation rule covering it
  // (Excel's "Clear All" on a selection). Dropped to undefined when none
  // remain so the no-validation path stays byte-identical.
  const clearValidationForSelection = useCallback((): boolean => {
    const authority = dataValidationMutationCommandRef.current;
    const liveSelection = readPageSetupSelectionSnapshot();
    const liveWorkbook = liveSheetsRef.current;
    const liveIndex = liveActiveSheetIndexRef.current;
    const liveSheet = liveWorkbook[liveIndex];
    if (
      !authority ||
      !liveSelection ||
      !liveSheet ||
      !spreadsheetDataValidationWriterReceiptMatches({
        expectedControlId: "XLSCTL.DATA.VALIDATION.MRAI.APPLY",
        receiptControlId: authority.controlId,
        sourceWorkbook: authority.workbook,
        sourceActiveSheetIndex: authority.activeSheetIndex,
        sourceSheet: authority.sheet,
        sourceSheetId: authority.sheetId,
        sourceSelectionEpoch: authority.selection.epoch,
        sourceRange: authority.selection.range,
        liveWorkbook,
        liveActiveSheetIndex: liveIndex,
        liveSheet,
        liveSheetId: liveSheet.id,
        liveSelectionEpoch: liveSelection.epoch,
        liveRange: liveSelection.range,
      }) ||
      isSheetProtectionEffectiveRef.current(liveSheet) ||
      rangeHasLockedCellRef.current(liveSheet, liveSelection.range)
    ) {
      return false;
    }
    const range: ValidationRange = { ...authority.selection.range };
    const liveValidations = normalizeRules(liveSheet.validations);
    // "Apply these changes to all other cells with the same settings" holds
    // for Clear All too: every rule carrying the opened cell's settings goes
    // with the selection's (B514), as the same checkbox with Any value does.
    const same =
      valApplyToSame && validationOpenedRule
        ? liveValidations.filter(sameValidationSettingsAs(validationOpenedRule))
        : [];
    if (!rangeHasRule(liveValidations, range) && same.length === 0) return false;
    if (
      spreadsheetNativeWriteRoute === "patchback" &&
      !spreadsheetDataValidationPatchbackSelectionSupported({
        selectionRange: range,
        rules: liveValidations,
      })
    ) {
      return false;
    }
    // Only the selected cells lose their rule; the rest of it stays.
    const rest = liveValidations.filter((candidate) => !same.includes(candidate));
    const next = rangeHasRule(rest, range)
      ? validationsAfterClear(rest, range, shiftFormula)
      : rest;
    if (!next) return false;
    const nextSheet: SheetModel = { ...liveSheet };
    if (next.length > 0) nextSheet.validations = withUntouchedRulesAsStored(liveSheet.validations, next) as typeof nextSheet.validations;
    else delete nextSheet.validations;
    replaceActiveSheet(nextSheet);
    return true;
  }, [
    readPageSetupSelectionSnapshot,
    replaceActiveSheet,
    spreadsheetNativeWriteRoute,
    valApplyToSame,
    validationOpenedRule,
  ]);
  return {
    clearValidationForSelection,
    defineValidationForSelection,
    validationPatchbackBlockerMessage,
  };
}
