"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The browser side of `POST /api/v1/office/import-projection`.
 *
 * Word, Excel and PowerPoint files are read on the server; the browser sends
 * the file and receives the content its caller consumes:
 *
 * - `document`: the nodes a document inserts (a guest's insert, or a document
 *   with no notebook to hold the file as its own page);
 * - `spreadsheet`: the workbook that replaces the current one;
 * - `xlsx_data_sources`: a workbook's external-data inventory.
 *
 * A signed-in caller asks with its session and CSRF token. The guest editor
 * has no session (the CSRF bootstrap refuses it), so it asks again without
 * one; the server serves an anonymous caller from an allowed origin under a
 * per-IP limit. With no connection the request fails as
 * {@link OfficeImportConnectionRequiredError}: there is no offline reader.
 */

import { apiFetchRaw, isApiRequestError } from "@mrnote/core/web/lib/api.ts";

export const OFFICE_IMPORT_PROJECTION_PATH = "/api/v1/office/import-projection";

export type OfficeImportSourceKind = "docx" | "xlsx" | "pptx";
export type OfficeImportTarget = "document" | "spreadsheet" | "presentation" | "xlsx_data_sources";

export interface OfficeImportLossRecord {
  readonly code: string;
  readonly omitted: number;
}

export interface OfficeDocumentImportResponse {
  readonly target: "document";
  readonly source_kind: OfficeImportSourceKind;
  readonly title: string;
  readonly nodes: Array<Record<string, unknown>>;
  readonly doc_attrs: Record<string, unknown> | null;
  readonly import_losses: readonly OfficeImportLossRecord[];
}

export interface OfficeSpreadsheetImportResponse {
  readonly target: "spreadsheet";
  readonly source_kind: "xlsx";
  readonly title: string;
  readonly content: Record<string, unknown>;
  readonly sheet_count: number;
  readonly cell_count: number;
  readonly sheet_names: readonly string[];
  readonly import_losses: readonly OfficeImportLossRecord[];
  /** The whole inventory when the workbook has external data, else null. */
  readonly data_source_inventory?: OfficeDataSourceInventoryRecord | null;
}

/** A deck page made of a .pptx, as the signed-in import makes it (no package). */
export interface OfficePresentationImportResponse {
  readonly target: "presentation";
  readonly source_kind: "pptx";
  readonly title: string;
  readonly content: Record<string, unknown>;
  readonly slide_count: number;
  readonly import_losses: readonly OfficeImportLossRecord[];
}

/** One connection of xl/connections.xml, as Connection Properties shows it. */
export interface OfficeConnectionRecord {
  readonly id: string;
  readonly name: string;
  readonly description: string;
  readonly kind: string;
  readonly power_query: boolean;
  readonly background_refresh: boolean;
  readonly refresh_interval: number;
  readonly refresh_on_load: boolean;
  readonly remove_data_on_save: boolean;
  readonly exclude_from_refresh_all: boolean;
  readonly save_password: boolean;
  readonly connection_file: string;
  readonly always_use_connection_file: boolean;
  readonly connection_string: string;
  readonly command_text: string;
  readonly command_type: number | null;
  readonly used_in: ReadonlyArray<{
    readonly sheet_name: string;
    readonly table_name: string;
    readonly location: string;
  }>;
}

/** One Power Query query of the workbook's DataMashup (query_mashup.py). */
export interface OfficeWorkbookQueryRecord {
  readonly name: string;
  readonly formula: string;
  readonly load: "table" | "connection_only" | "data_model";
  readonly loads_to_data_model: boolean;
  readonly private: boolean;
  readonly result_type: string | null;
  readonly last_refreshed: string | null;
  readonly fill_count: number | null;
  readonly fill_error_count: number | null;
  readonly query_id: string | null;
  readonly connection_id: string | null;
  readonly connection_name: string | null;
  readonly description: string;
  readonly sheet_name: string | null;
  readonly table_name: string | null;
  readonly location: string | null;
  readonly columns: readonly string[];
  readonly source_kind: "local" | "external" | "none";
  readonly functions: readonly string[];
  readonly workbook_sources: readonly string[];
  readonly query_references: readonly string[];
  readonly external_functions: readonly string[];
  readonly culture: string;
  readonly mashup_part: string;
  readonly mashup_sha256: string;
}

export interface OfficeDataSourceInventoryRecord {
  readonly connection_parts: number;
  readonly query_table_parts: number;
  readonly workbook_link_parts: number;
  readonly linked_data_parts: number;
  readonly workbook_links: ReadonlyArray<{
    readonly id: string;
    readonly part_name: string;
    readonly display_source: string;
    readonly source_kind: "web" | "file" | "unknown";
  }>;
  readonly connections?: readonly OfficeConnectionRecord[];
  readonly queries?: readonly OfficeWorkbookQueryRecord[];
}

export interface OfficeDataSourcesImportResponse extends OfficeDataSourceInventoryRecord {
  readonly target: "xlsx_data_sources";
}

type ResponseFor<T extends OfficeImportTarget> = T extends "document"
  ? OfficeDocumentImportResponse
  : T extends "spreadsheet"
    ? OfficeSpreadsheetImportResponse
    : T extends "presentation"
      ? OfficePresentationImportResponse
      : OfficeDataSourcesImportResponse;

/** No connection to the server, which is the only Office reader there is. */
export class OfficeImportConnectionRequiredError extends Error {
  readonly code = "office_import_connection_required";

  constructor() {
    super("office_import_connection_required");
    this.name = "OfficeImportConnectionRequiredError";
  }
}

/** The server read the request and refused it, in its own code. */
export class OfficeImportRefusedError extends Error {
  constructor(
    readonly code: string,
    readonly status: number,
    message: string,
    readonly details: Record<string, unknown> = {},
  ) {
    super(message || code);
    this.name = "OfficeImportRefusedError";
  }
}

export function isOfficeImportConnectionRequired(error: unknown): boolean {
  return error instanceof OfficeImportConnectionRequiredError;
}

/** The refusals that mean "this file is too large to be read here". */
export const OFFICE_IMPORT_TOO_LARGE_CODES: ReadonlySet<string> = new Set([
  "editable_source_too_large",
  "office_import_too_large_for_page",
  "payload_too_large",
  "file_too_large",
]);

function aborted(error: unknown): boolean {
  return error instanceof DOMException && error.name === "AbortError";
}

function mapError(error: unknown): unknown {
  if (aborted(error)) return error;
  if (isApiRequestError(error)) {
    if (error.status === 0 || error.code === "network_unreachable") {
      return new OfficeImportConnectionRequiredError();
    }
    return new OfficeImportRefusedError(
      error.code ?? `http_${error.status}`,
      error.status,
      error.message,
      error.details ?? {},
    );
  }
  if (error instanceof TypeError) return new OfficeImportConnectionRequiredError();
  return error;
}

/** The CSRF bootstrap refused (no session), or the session itself did. */
function noSession(error: unknown): boolean {
  if (aborted(error)) return false;
  if (isApiRequestError(error)) {
    return error.status === 401 || (error.status === 403 && /csrf/i.test(error.code ?? ""));
  }
  return error instanceof Error && !(error instanceof TypeError);
}

interface Ask {
  readonly file: Blob;
  readonly mediaType: string;
  readonly headers: Record<string, string>;
}

function ask(request: Ask, signedIn: boolean, signal?: AbortSignal): Promise<Response> {
  // The body is the file itself (the guest editor has no session, and every
  // multipart route sits behind the authenticated-upload gate); what to read
  // it as rides in headers, never in the URL.
  return apiFetchRaw(
    OFFICE_IMPORT_PROJECTION_PATH,
    {
      method: "POST",
      body: request.file,
      headers: request.headers,
      ...(signal ? { signal } : {}),
    },
    {
      contentType: request.mediaType,
      // An import never bounces the page to the login screen.
      suppressUnauthorizedHandling: true,
      ...(signedIn ? {} : { requireCsrf: false }),
    },
  );
}

const MEDIA_TYPES: Record<OfficeImportSourceKind, string> = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
};

function declaredMediaType(file: Blob, filename: string, kind: OfficeImportSourceKind): string {
  const lower = filename.toLowerCase();
  if (lower.endsWith(".docm")) return "application/vnd.ms-word.document.macroenabled.12";
  if (lower.endsWith(".pptm")) return "application/vnd.ms-powerpoint.presentation.macroenabled.12";
  if (lower.endsWith(".potm")) return "application/vnd.ms-powerpoint.template.macroenabled.12";
  if (lower.endsWith(".ppsm")) return "application/vnd.ms-powerpoint.slideshow.macroenabled.12";
  if (lower.endsWith(".potx")) return "application/vnd.openxmlformats-officedocument.presentationml.template";
  if (lower.endsWith(".ppsx")) return "application/vnd.openxmlformats-officedocument.presentationml.slideshow";
  if (lower.endsWith(".xls")) return "application/vnd.ms-excel";
  if (lower.endsWith(".xlsm")) return "application/vnd.ms-excel.sheet.macroenabled.12";
  if (lower.endsWith(".xltx")) return "application/vnd.openxmlformats-officedocument.spreadsheetml.template";
  if (lower.endsWith(".xltm")) return "application/vnd.ms-excel.template.macroenabled.12";
  // Excel's other workbook files: the server names each by its extension
  // (a browser gives .numbers or .et no type, or a generic one).
  if (/\.(xlt|xlsb|ods|fods|et|ett|slk|dif|dbf|numbers|xml)$/.test(lower)) return "application/octet-stream";
  const type = (file.type || "").split(";", 1)[0].trim().toLowerCase();
  return type && type !== "application/octet-stream" ? type : MEDIA_TYPES[kind];
}

export async function requestOfficeImportProjection<T extends OfficeImportTarget>(
  file: Blob,
  options: {
    readonly sourceKind: OfficeImportSourceKind;
    readonly target: T;
    readonly filename: string;
    readonly signal?: AbortSignal;
    /**
     * `custody`: a workbook whose queries, charts or SmartArt need its original
     * package is read anyway, those parts said as losses (the homepage's
     * signed-out editor, which has no page to replace).
     */
    readonly acceptLosses?: "custody";
    /**
     * `first-pages`: a long Word file's first pages only, while its page is
     * imported (lib/document-import/docx-first-pages); the answer then says
     * `preview: "first_pages"`, or `preview: "none"` with no nodes.
     */
    readonly preview?: "first-pages";
  },
): Promise<ResponseFor<T>> {
  if (typeof navigator !== "undefined" && navigator.onLine === false) {
    throw new OfficeImportConnectionRequiredError();
  }
  const request: Ask = {
    file,
    mediaType: declaredMediaType(file, options.filename, options.sourceKind),
    headers: {
      "X-Office-Source-Kind": options.sourceKind,
      "X-Office-Import-Target": options.target,
      "X-File-Name": encodeURIComponent(options.filename),
      ...(options.acceptLosses ? { "X-Office-Import-Accept-Losses": options.acceptLosses } : {}),
      ...(options.preview ? { "X-Office-Import-Preview": options.preview } : {}),
    },
  };
  let raw: Response;
  try {
    try {
      raw = await ask(request, true, options.signal);
    } catch (error) {
      if (!noSession(error)) throw error;
      raw = await ask(request, false, options.signal);
    }
  } catch (error) {
    throw mapError(error);
  }
  let payload: unknown;
  try {
    payload = await raw.json();
  } catch (error) {
    if (aborted(error)) throw error;
    throw new OfficeImportRefusedError("malformed_response", raw.status, "malformed response");
  }
  if (!payload || typeof payload !== "object" || (payload as { target?: unknown }).target !== options.target) {
    throw new OfficeImportRefusedError("malformed_response", raw.status, "malformed response");
  }
  return payload as ResponseFor<T>;
}

/** The source kind a file is read as, from its name or media type. */
export function officeImportSourceKind(file: { name: string; type?: string }): OfficeImportSourceKind | null {
  const name = file.name.toLowerCase();
  if (name.endsWith(".docx") || name.endsWith(".docm")) return "docx";
  if (/\.(xlsx|xls|xlsm|xltx|xltm|xlt|xlsb|ods|fods|et|ett|slk|dif|dbf|numbers)$/.test(name)) return "xlsx";
  if (name.endsWith(".pptx")) return "pptx";
  const type = (file.type || "").split(";", 1)[0].toLowerCase();
  if (
    type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document" ||
    type === "application/vnd.ms-word.document.macroenabled.12"
  )
    return "docx";
  if (type === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") return "xlsx";
  if (type === "application/vnd.openxmlformats-officedocument.presentationml.presentation") return "pptx";
  return null;
}
