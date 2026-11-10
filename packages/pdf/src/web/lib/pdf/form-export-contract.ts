/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Which file convention produced an entry. Absent means an exact typed value
 * from an in-app caller (the viewer's live overrides), read like format 2.
 */
export type PdfFormDataDialect =
  | "csv-v1"
  | "csv-v2"
  | "xfdf-unversioned"
  | "xfdf-v2";

/** A single form field entry: its fully-qualified name and typed value. */
export interface PdfFormFieldValue {
  /** Exact fully-qualified field name; identity is never trimmed. */
  name: string;
  /** Scalar value, or a ", "-joined display string when `values` is set. */
  value: string;
  /**
   * Typed list value (choice fields). When present it is authoritative and
   * may be empty; `value` is then display-only and never split.
   */
  values?: ReadonlyArray<string>;
  /** Set by the parsers so legacy files keep their original meaning. */
  dialect?: PdfFormDataDialect;
}

export type PdfFormDataFormat = "csv" | "xfdf";

export interface PdfFormDataFileDescriptor {
  readonly name: string;
  readonly type?: string;
}

/**
 * Resolve a local form-data import through one parser-owned catalog boundary.
 * The picker hint is not authority: the bounded CSV/XFDF parser still validates
 * the full payload before any form mutation or output is allowed.
 */
export function resolvePdfFormDataFormatForFile(
  file: PdfFormDataFileDescriptor,
): PdfFormDataFormat | null {
  const normalizedName = file.name.trim().toLowerCase();
  const dot = normalizedName.lastIndexOf(".");
  const extension = dot >= 0 ? normalizedName.slice(dot) : "";
  if (extension === ".xfdf") return "xfdf";
  if (extension === ".csv") return "csv";

  const mediaType = (file.type || "").split(";", 1)[0].trim().toLowerCase();
  if (mediaType === "application/vnd.adobe.xfdf") return "xfdf";
  if (mediaType === "text/csv") return "csv";
  return null;
}

export type PdfFormRewriteBoundaryCode =
  | "aggregate_limit"
  | "apply_failed"
  | "csv_columns_limit"
  | "csv_malformed"
  | "csv_rows_limit"
  | "csv_unsupported"
  | "field_limit"
  | "field_unreadable"
  | "name_limit"
  | "output_limit"
  | "override_rejected"
  | "plan_mismatch"
  | "reset_failed"
  | "value_limit"
  | "xfdf_malformed"
  | "xfdf_unsupported";

/**
 * Why one field made an export incomplete. Read failures never degrade to an
 * empty value, and a rejected current value never falls back to the source.
 */
export type PdfFormDataIssueReason =
  | PdfFormImportSkipReason
  | "form_unreadable"
  | "name_unreadable"
  | "text_unrepresentable"
  | "value_unreadable";

export interface PdfFormDataIssue {
  /** Exact field name, or null when the name itself could not be read. */
  readonly name: string | null;
  readonly reason: PdfFormDataIssueReason;
}

/** Stable fail-closed error returned by the bounded form data primitives. */
export class PdfFormRewriteBoundaryError extends Error {
  readonly code: PdfFormRewriteBoundaryCode;
  /** Per-field detail for export refusals; empty for structural limits. */
  readonly issues: ReadonlyArray<PdfFormDataIssue>;

  constructor(
    code: PdfFormRewriteBoundaryCode,
    issues: ReadonlyArray<PdfFormDataIssue> = [],
  ) {
    super(`PDF form rewrite rejected: ${code}`);
    this.name = "PdfFormRewriteBoundaryError";
    this.code = code;
    this.issues = Object.freeze(
      issues.map((issue) => Object.freeze({ ...issue })),
    );
  }
}

export type PdfFormImportSkipReason =
  | "field_not_found"
  | "invalid_value"
  | "read_only"
  | "required_empty"
  | "unsupported_field";
