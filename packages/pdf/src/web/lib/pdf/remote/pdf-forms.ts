/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type * as Create from "../form-field-create";
import type * as TextEdit from "../form-field-edit";
import type * as CheckboxEdit from "../form-checkbox-field-edit";
import type * as DropdownEdit from "../form-dropdown-field-edit";
import type * as ListBoxEdit from "../form-listbox-field-edit";
import type * as SignatureEdit from "../form-signature-field-edit";
import type * as ActionButtonEdit from "../form-action-button-field-edit";
import type * as Delete from "../form-field-delete";
import type * as FormExport from "../form-export";
import {
  PdfFormRewriteBoundaryError,
  type PdfFormDataIssue,
  type PdfFormRewriteBoundaryCode,
} from "../form-export-contract";
import { callPdfDocumentSourceless } from "./pdf-export";

export * from "../form-export-contract";
export * from "../form-field-contract";

// eslint-disable-next-line @typescript-eslint/no-explicit-any -- any planner signature
type Fn = (...args: any[]) => unknown;
type Remote<F extends Fn> = (...args: Parameters<F>) => Promise<Awaited<ReturnType<F>>>;

interface IssuedPlan {
  readonly planFn: string;
  readonly planArgs: readonly unknown[];
  /** createFormFieldsPlan: its inputs, for the review's freshness check. */
  readonly inputKey?: string;
}

const issued = new WeakMap<object, IssuedPlan>();

function stableJson(value: unknown): string {
  if (value instanceof Map) {
    return `M${stableJson([...value].sort(([a], [b]) => String(a).localeCompare(String(b))))}`;
  }
  if (value instanceof Set) return `S${stableJson([...value])}`;
  if (Array.isArray(value)) return `[${value.map(stableJson).join(",")}]`;
  if (value && typeof value === "object") {
    const record = value as Record<string, unknown>;
    return `{${Object.keys(record).sort().filter((key) => record[key] !== undefined && typeof record[key] !== "function")
      .map((key) => `${JSON.stringify(key)}:${stableJson(record[key])}`).join(",")}}`;
  }
  return JSON.stringify(value ?? null);
}

function createFormFieldsInputKey(candidates: unknown, options: { pageSizes?: unknown } | undefined): string {
  return stableJson([candidates, options?.pageSizes ?? null]);
}

async function plan(fn: string, args: readonly unknown[]): Promise<unknown> {
  return callPdfDocumentSourceless("form.plan", [fn, [...args]], { fn: "formPlan" });
}

function planner<F extends Fn>(fn: string, remember: boolean): Remote<F> {
  return (async (...args: Parameters<F>) => {
    const value = await plan(fn, args);
    if (remember && value && typeof value === "object") {
      issued.set(value as object, {
        planFn: fn,
        planArgs: args,
        ...(fn === "createFormFieldsPlan"
          ? { inputKey: createFormFieldsInputKey(args[0], args[1] as { pageSizes?: unknown }) }
          : {}),
      });
    }
    return value;
  }) as Remote<F>;
}

// -- candidates (the review list) ------------------------------------------------
export const createAuthoredCheckBoxCandidate =
  planner<typeof Create.createAuthoredCheckBoxCandidate>("createAuthoredCheckBoxCandidate", false);
export const createAuthoredDropdownCandidate =
  planner<typeof Create.createAuthoredDropdownCandidate>("createAuthoredDropdownCandidate", false);
export const createAuthoredGoToPageButtonCandidate =
  planner<typeof Create.createAuthoredGoToPageButtonCandidate>("createAuthoredGoToPageButtonCandidate", false);
export const createAuthoredListBoxCandidate =
  planner<typeof Create.createAuthoredListBoxCandidate>("createAuthoredListBoxCandidate", false);
export const createAuthoredNavigationButtonCandidate =
  planner<typeof Create.createAuthoredNavigationButtonCandidate>("createAuthoredNavigationButtonCandidate", false);
export const createAuthoredPrintButtonCandidate =
  planner<typeof Create.createAuthoredPrintButtonCandidate>("createAuthoredPrintButtonCandidate", false);
export const createAuthoredRadioGroupCandidate =
  planner<typeof Create.createAuthoredRadioGroupCandidate>("createAuthoredRadioGroupCandidate", false);
export const createAuthoredResetButtonCandidate =
  planner<typeof Create.createAuthoredResetButtonCandidate>("createAuthoredResetButtonCandidate", false);
export const createAuthoredSaveButtonCandidate =
  planner<typeof Create.createAuthoredSaveButtonCandidate>("createAuthoredSaveButtonCandidate", false);
export const createAuthoredSignatureFieldCandidate =
  planner<typeof Create.createAuthoredSignatureFieldCandidate>("createAuthoredSignatureFieldCandidate", false);
export const createAuthoredTextFieldCandidate =
  planner<typeof Create.createAuthoredTextFieldCandidate>("createAuthoredTextFieldCandidate", false);
export const createAuthoredViewButtonCandidate =
  planner<typeof Create.createAuthoredViewButtonCandidate>("createAuthoredViewButtonCandidate", false);
export const createAuthoredWebLinkButtonCandidate =
  planner<typeof Create.createAuthoredWebLinkButtonCandidate>("createAuthoredWebLinkButtonCandidate", false);

// -- plans ---------------------------------------------------------------------
export const createFormFieldsPlan =
  planner<typeof Create.createFormFieldsPlan>("createFormFieldsPlan", true);
export const createExistingPdfTextFieldEditPlan =
  planner<typeof TextEdit.createExistingPdfTextFieldEditPlan>("createExistingPdfTextFieldEditPlan", true);
export const createExistingPdfCheckboxFieldEditPlan =
  planner<typeof CheckboxEdit.createExistingPdfCheckboxFieldEditPlan>("createExistingPdfCheckboxFieldEditPlan", true);
export const createExistingPdfDropdownFieldEditPlan =
  planner<typeof DropdownEdit.createExistingPdfDropdownFieldEditPlan>("createExistingPdfDropdownFieldEditPlan", true);
export const createExistingPdfListBoxFieldEditPlan =
  planner<typeof ListBoxEdit.createExistingPdfListBoxFieldEditPlan>("createExistingPdfListBoxFieldEditPlan", true);
export const createExistingPdfSignatureFieldEditPlan =
  planner<typeof SignatureEdit.createExistingPdfSignatureFieldEditPlan>("createExistingPdfSignatureFieldEditPlan", true);
export const createExistingPdfActionButtonFieldEditPlan =
  planner<typeof ActionButtonEdit.createExistingPdfActionButtonFieldEditPlan>(
    "createExistingPdfActionButtonFieldEditPlan", true);
export const createExistingPdfFormFieldDeletePlan =
  planner<typeof Delete.createExistingPdfFormFieldDeletePlan>("createExistingPdfFormFieldDeletePlan", true);

/**
 * The review is still the one this plan was made from: the same plan (one
 * this module received) and the same candidates and page sizes. Planning is
 * deterministic, so equal inputs are the server's equal plan.
 */
export function matchesCreateFormFieldsPlan(
  plan: Create.CreateFormFieldsPlan,
  candidates: ReadonlyArray<Create.CreateFormFieldCandidate>,
  options: Create.CreateFormFieldsOptions,
): boolean {
  const record = issued.get(plan);
  return Boolean(record && record.planFn === "createFormFieldsPlan" &&
    record.inputKey === createFormFieldsInputKey(candidates, options));
}

function jobOf(planValue: { digestParts: readonly string[] }) {
  const record = issued.get(planValue);
  if (!record) throw new Error("PDF form plan was not issued here");
  return { planFn: record.planFn, planArgs: record.planArgs, digestParts: planValue.digestParts };
}

type FormWriteAnswer = Awaited<ReturnType<typeof import("../form-writers-server").formWrite>>;

/** Apply and verify the plan in memory on the reader's PDF (before consent). */
export async function preflightPdfFormPlan(
  sourceBytes: Uint8Array,
  planValue: { digestParts: readonly string[] },
): Promise<FormWriteAnswer> {
  return callPdfDocumentSourceless("form", [sourceBytes, jobOf(planValue), "preflight"],
    { source: sourceBytes, fn: "formWrite" }) as Promise<FormWriteAnswer>;
}

/** Write the plan into the reader's PDF; answers the reopened, verified file. */
export async function writePdfFormPlan(
  sourceBytes: Uint8Array,
  planValue: { digestParts: readonly string[] },
): Promise<FormWriteAnswer & { bytes: Uint8Array }> {
  const answer = await callPdfDocumentSourceless("form", [sourceBytes, jobOf(planValue), "write"],
    { source: sourceBytes, fn: "formWrite" }) as FormWriteAnswer;
  if (!(answer.bytes instanceof Uint8Array)) throw new Error("PDF form write answered no file");
  return answer as FormWriteAnswer & { bytes: Uint8Array };
}

// -- form data (CSV / XFDF) --------------------------------------------------------
function rebuildBoundary(error: { name: string; code?: string; issues?: unknown }): Error | null {
  if (error.name !== "PdfFormRewriteBoundaryError" || !error.code) return null;
  const issues = Array.isArray(error.issues) ? (error.issues as PdfFormDataIssue[]) : [];
  return new PdfFormRewriteBoundaryError(error.code as PdfFormRewriteBoundaryCode, issues);
}

function asBytes(bytes: Uint8Array | ArrayBuffer): Uint8Array {
  return bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
}

export function exportPdfFormData(
  ...args: Parameters<typeof FormExport.exportPdfFormData>
): ReturnType<typeof FormExport.exportPdfFormData> {
  const source = asBytes(args[0]);
  return callPdfDocumentSourceless("form", [source, ...args.slice(1)],
    { source, fn: "exportPdfFormData", rebuild: rebuildBoundary }) as ReturnType<typeof FormExport.exportPdfFormData>;
}

export function preflightPdfFormReset(
  ...args: Parameters<typeof FormExport.preflightPdfFormReset>
): ReturnType<typeof FormExport.preflightPdfFormReset> {
  const source = asBytes(args[0]);
  return callPdfDocumentSourceless("form", [source, ...args.slice(1)],
    { source, fn: "preflightPdfFormReset", rebuild: rebuildBoundary }) as ReturnType<typeof FormExport.preflightPdfFormReset>;
}

export function preflightPdfFormDataImport(
  ...args: Parameters<typeof FormExport.preflightPdfFormDataImport>
): ReturnType<typeof FormExport.preflightPdfFormDataImport> {
  const source = asBytes(args[0]);
  return callPdfDocumentSourceless("form", [source, ...args.slice(1)],
    { source, fn: "preflightPdfFormDataImport", rebuild: rebuildBoundary }) as ReturnType<
      typeof FormExport.preflightPdfFormDataImport>;
}
