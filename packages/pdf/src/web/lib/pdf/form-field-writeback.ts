/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  ACROFORM_MAX_OPTIONS,
  ACROFORM_MAX_TEXT,
  normalizeFieldName,
  type AcroFormValue,
} from "@mrnote/core/web/lib/pdf/acroform.ts";

export interface NativePdfFormFieldEdit {
  id: string;
  value: AcroFormValue;
  label?: string;
  /** 1-based PDF page containing the edited widget. */
  page?: number;
  




  displayValue?: string | string[];
}

export interface PersistedPdfFormField extends NativePdfFormFieldEdit {
  updatedAt?: string;
}

export interface NativePdfFormWriteAuthority {
  readonly generation: string;
  readonly writable: boolean;
  readonly reason?: string;
}

export interface NativePdfFormFieldMutation {
  readonly generation: string;
  /** One focus interaction or one discrete control activation. */
  readonly gestureId: string;
  readonly fields: readonly NativePdfFormFieldEdit[];
}

export interface NativePdfFormPendingGesture extends NativePdfFormFieldMutation {
  readonly closed: boolean;
  /** Its fields as shown before it changed them. */
  readonly previous?: readonly NativePdfFormFieldEdit[];
}






export interface NativePdfFormPendingSnapshot {
  readonly generation: string;
  readonly saving: boolean;
  readonly gestures: readonly NativePdfFormPendingGesture[];
}

function normalizeFormFieldPage(page: unknown): number | undefined {
  if (typeof page !== "number" || !Number.isInteger(page) || page < 1) {
    return undefined;
  }
  return page;
}

function normalizeFormFieldValue(value: AcroFormValue): AcroFormValue {
  if (typeof value === "string") return value.slice(0, ACROFORM_MAX_TEXT);
  if (!Array.isArray(value)) return value;
  const out: string[] = [];
  const seen = new Set<string>();
  for (const item of value) {
    const normalized = item.slice(0, ACROFORM_MAX_TEXT);
    if (seen.has(normalized)) continue;
    seen.add(normalized);
    out.push(normalized);
    if (out.length >= ACROFORM_MAX_OPTIONS) break;
  }
  return out;
}

/**
 * Upsert one viewer form edit into the native page model without reordering
 * existing fields. Explicit clears (`""` / `false` / `[]`) are data and must
 * survive: the PDF patchback bridge uses them to clear a value already baked
 * into the original AcroForm.
 */
export function upsertPdfFormField(
  fields: readonly PersistedPdfFormField[],
  edit: NativePdfFormFieldEdit,
  updatedAt = new Date().toISOString(),
): PersistedPdfFormField[] {
  const id = normalizeFieldName(edit.id);
  if (!id) return [...fields];

  const nextField: PersistedPdfFormField = {
    id,
    value: normalizeFormFieldValue(edit.value),
    ...(typeof edit.label === "string" && edit.label
      ? { label: edit.label.slice(0, ACROFORM_MAX_TEXT) }
      : {}),
    ...(normalizeFormFieldPage(edit.page) !== undefined
      ? { page: normalizeFormFieldPage(edit.page) }
      : {}),
    ...(typeof edit.displayValue === "string"
      ? { displayValue: edit.displayValue.slice(0, ACROFORM_MAX_TEXT) }
      : Array.isArray(edit.displayValue) &&
          edit.displayValue.every((item) => typeof item === "string")
        ? {
            displayValue: edit.displayValue
              .slice(0, ACROFORM_MAX_OPTIONS)
              .map((item) => item.slice(0, ACROFORM_MAX_TEXT)),
          }
        : {}),
    updatedAt,
  };
  const index = fields.findIndex((field) => field.id === id);
  if (index < 0) return [...fields, nextField];
  return fields.map((field, fieldIndex) =>
    fieldIndex === index ? { ...field, ...nextField } : field,
  );
}

/**
 * Apply one user gesture's complete AcroForm value set with a shared timestamp.
 * A calculation pass can update several dependent fields; keeping the batch
 * ordered and timestamp-identical lets the native host save and undo it as one
 * receipt-bound transaction instead of one write per calculated field.
 */
export function upsertPdfFormFields(
  fields: readonly PersistedPdfFormField[],
  edits: readonly NativePdfFormFieldEdit[],
  updatedAt = new Date().toISOString(),
): PersistedPdfFormField[] {
  return edits.reduce<PersistedPdfFormField[]>(
    (current, edit) => upsertPdfFormField(current, edit, updatedAt),
    [...fields],
  );
}

function formValuesEqual(
  left: NativePdfFormFieldEdit["value"],
  right: NativePdfFormFieldEdit["value"],
): boolean {
  if (!Array.isArray(left) || !Array.isArray(right)) return left === right;
  return (
    left.length === right.length &&
    left.every((value, index) => value === right[index])
  );
}

/** Semantic equality intentionally ignores `updatedAt`. */
export function pdfFormFieldsEqual(
  left: readonly PersistedPdfFormField[],
  right: readonly PersistedPdfFormField[],
): boolean {
  return (
    left.length === right.length &&
    left.every((field, index) => {
      const candidate = right[index];
      return Boolean(
        candidate &&
        field.id === candidate.id &&
        field.label === candidate.label &&
        field.page === candidate.page &&
        formValuesEqual(field.value, candidate.value),
      );
    })
  );
}
