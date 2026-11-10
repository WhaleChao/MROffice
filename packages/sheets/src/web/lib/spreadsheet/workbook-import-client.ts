"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * File > Import of an .xlsx into the CURRENT workbook, and the workbook's
 * external-data inventory, both read on the server.
 *
 * The browser keeps no Excel reader: the workbook is read by the persisted
 * import's own projection (`POST /api/v1/office/import-projection`, target
 * `spreadsheet`), which also refuses what this page could not carry without
 * the file's package -- native charts, queries, a SmartArt diagram it can only
 * preserve, more cells than a page stores -- so the caller can offer to open
 * the file as a new, source-backed workbook instead. Losses still need the
 * reader's consent: the first call throws {@link SpreadsheetWorkbookLossError}
 * with the manifest, and the consented call reuses the answer it already has.
 */

import type { NativeSpreadsheetContent } from "@mrnote/core/web/lib/native-page-schemas.ts";
import {
  OfficeImportRefusedError,
  requestOfficeImportProjection,
  type OfficeDataSourceInventoryRecord,
  type OfficeSpreadsheetImportResponse,
} from "@mrnote/docs/web/lib/document-import/officeImportProjectionClient.ts";
import {
  describeSpreadsheetImportLosses,
  type SpreadsheetImportLossLocale,
} from "@mrnote/core/web/lib/spreadsheet/workbook-import-losses.ts";

export interface SpreadsheetWorkbookLossManifest {
  status: "lossless" | "lossy";
  entries: { code: string; affected: number; limit?: number; detail?: string }[];
}

export interface SpreadsheetWorkbookImportPreview {
  sheetCount: number;
  cellCount: number;
  sheetNames: string[];
}

export interface SpreadsheetWorkbookImportResult {
  content: NativeSpreadsheetContent;
  title: string;
  lossManifest: SpreadsheetWorkbookLossManifest;
  sheetCount: number;
  cellCount: number;
  /**
   * The workbook's whole data-source inventory (workbook links' sources and
   * each connection's Connection Properties) when it has external data: the
   * replaced page keeps no source package to read it from later.
   */
  dataSourceInventory?: SpreadsheetDataSourceInventory;
}

/**
 * Bounded inventory of external-data package parts.
 *
 * The counts expose no provider payloads or cached values. They let the Data
 * tab distinguish an empty workbook from an imported workbook whose native
 * data objects are being preserved/read-only while no provider-backed
 * refresh authority exists. `connections` adds what Excel's (read-only)
 * Connection Properties shows of each connection, any password in a
 * connection string masked by the server's reader
 * (apps/api app/services/file_adapters/xlsx/connection_inventory.py).
 */
export interface SpreadsheetDataSourceInventory {
  readonly connectionParts: number;
  readonly queryTableParts: number;
  readonly workbookLinkParts: number;
  readonly linkedDataParts: number;
  readonly workbookLinks: readonly SpreadsheetWorkbookLinkInventoryEntry[];
  /** Each connection, for its read-only Connection / Query Properties. */
  readonly connections?: readonly SpreadsheetConnectionInventoryEntry[];
  /** The workbook's Power Query queries (Queries & Connections). */
  readonly queries?: readonly SpreadsheetWorkbookQuery[];
}

/** A Power Query query as Excel's Queries & Connections pane lists it. */
export interface SpreadsheetWorkbookQuery {
  readonly name: string;
  /** The query's M expression (the section member's right-hand side). */
  readonly formula: string;
  readonly load: "table" | "connection_only" | "data_model";
  readonly loadsToDataModel: boolean;
  readonly lastRefreshed: string | null;
  readonly rowCount: number | null;
  readonly errorCount: number | null;
  readonly connectionId: string | null;
  readonly description: string;
  readonly sheetName: string | null;
  readonly tableName: string | null;
  readonly location: string | null;
  readonly columns: readonly string[];
  /** "local": only this workbook's tables/ranges and its other queries. */
  readonly sourceKind: "local" | "external" | "none";
  readonly functions: readonly string[];
  readonly workbookSources: readonly string[];
  readonly queryReferences: readonly string[];
  readonly externalFunctions: readonly string[];
  /** Config/Package.xml Culture: number and date text conversions. */
  readonly culture: string;
}

export type SpreadsheetConnectionKind =
  | "odbc"
  | "dao"
  | "file"
  | "web"
  | "oledb"
  | "text"
  | "ado"
  | "dsp"
  | "unknown";

export interface SpreadsheetConnectionUse {
  readonly sheetName: string;
  readonly tableName: string;
  readonly location: string;
}

export interface SpreadsheetConnectionInventoryEntry {
  readonly id: string;
  readonly name: string;
  readonly description: string;
  readonly kind: SpreadsheetConnectionKind;
  /** An OLE DB connection to Power Query's Mashup provider. */
  readonly powerQuery: boolean;
  readonly backgroundRefresh: boolean;
  readonly refreshInterval: number;
  readonly refreshOnLoad: boolean;
  readonly removeDataOnSave: boolean;
  readonly excludeFromRefreshAll: boolean;
  readonly savePassword: boolean;
  readonly connectionFile: string;
  readonly alwaysUseConnectionFile: boolean;
  /** dbPr/@connection with any password masked, or the source of a text/web connection. */
  readonly connectionString: string;
  readonly commandText: string;
  /** ST_CmdType: 1 cube, 2 SQL, 3 table, 4 default, 5 list. */
  readonly commandType: number | null;
  readonly usedIn: readonly SpreadsheetConnectionUse[];
}

export interface SpreadsheetWorkbookLinkInventoryEntry {
  readonly id: string;
  readonly partName: string;
  readonly displaySource: string;
  readonly sourceKind: "web" | "file" | "unknown";
}

/** The workbook's queries need its package: open it as a new workbook. */
export class SpreadsheetQuerySourceCustodyRequiredError extends Error {
  constructor() {
    super("xlsx_query_source_custody_required");
  }
}

/** The workbook's native charts need its package: open it as a new workbook. */
export class SpreadsheetChartSourceCustodyRequiredError extends Error {
  constructor() {
    super("xlsx_chart_source_custody_required");
  }
}

/** A SmartArt diagram only its package carries: open it as a new workbook. */
export class SpreadsheetDiagramSourceCustodyRequiredError extends Error {
  constructor() {
    super("xlsx_diagram_source_custody_required");
  }
}

/**
 * The workbook holds more cells than a page can store, so importing it into
 * this page could never be saved. The server imports such a file as a
 * read-only preview of each sheet's top rows, which the caller offers instead.
 */
export class SpreadsheetWorkbookTooLargeError extends Error {
  readonly code = "spreadsheet_workbook_too_large_for_page";

  constructor(
    readonly sheetName: string,
    readonly cellCount: number,
  ) {
    super("spreadsheet_workbook_too_large_for_page");
    this.name = "SpreadsheetWorkbookTooLargeError";
  }
}

export class SpreadsheetWorkbookLossError extends Error {
  readonly code = "spreadsheet_workbook_import_lossy";

  constructor(
    readonly lossManifest: SpreadsheetWorkbookLossManifest,
    readonly preview?: SpreadsheetWorkbookImportPreview,
  ) {
    super("spreadsheet_workbook_import_lossy");
    this.name = "SpreadsheetWorkbookLossError";
  }
}

const answersByDefault = new WeakMap<File, Promise<OfficeSpreadsheetImportResponse>>();
const custodyAcceptedAnswers = new WeakMap<File, Promise<OfficeSpreadsheetImportResponse>>();

function readOnce(
  file: File,
  signal?: AbortSignal,
  acceptCustodyLosses = false,
): Promise<OfficeSpreadsheetImportResponse> {
  const answers = acceptCustodyLosses ? custodyAcceptedAnswers : answersByDefault;
  const cached = answers.get(file);
  if (cached) return cached;
  const pending = requestOfficeImportProjection(file, {
    sourceKind: "xlsx",
    target: "spreadsheet",
    filename: file.name,
    signal,
    ...(acceptCustodyLosses ? { acceptLosses: "custody" as const } : {}),
  });
  answers.set(file, pending);
  void pending.catch(() => {
    if (answers.get(file) === pending) answers.delete(file);
  });
  return pending;
}

function mapRefusal(error: unknown): unknown {
  if (!(error instanceof OfficeImportRefusedError)) return error;
  switch (error.code) {
    case "xlsx_query_source_custody_required":
      return new SpreadsheetQuerySourceCustodyRequiredError();
    case "xlsx_chart_source_custody_required":
      return new SpreadsheetChartSourceCustodyRequiredError();
    case "xlsx_diagram_source_custody_required":
      return new SpreadsheetDiagramSourceCustodyRequiredError();
    case "spreadsheet_workbook_too_large_for_page":
    case "office_import_too_large_for_page":
      return new SpreadsheetWorkbookTooLargeError(
        typeof error.details.sheet_name === "string" ? error.details.sheet_name : "",
        typeof error.details.cell_count === "number" ? error.details.cell_count : 0,
      );
    default:
      return error;
  }
}

/** Read an .xlsx into a workbook that replaces the current one. */
export async function importSpreadsheetWorkbook(
  file: File,
  options: {
    readonly allowLossy?: boolean;
    readonly signal?: AbortSignal;
    /** Read a workbook whose queries/charts/SmartArt need its package (no
     *  page is replaced: the homepage's signed-out editor). */
    readonly acceptSourceCustodyLosses?: boolean;
  } = {},
): Promise<SpreadsheetWorkbookImportResult> {
  let response: OfficeSpreadsheetImportResponse;
  try {
    response = await readOnce(file, options.signal, options.acceptSourceCustodyLosses === true);
  } catch (error) {
    throw mapRefusal(error);
  }
  const entries = response.import_losses.map((entry) => ({
    code: entry.code,
    affected: Math.max(1, entry.omitted),
  }));
  const lossManifest: SpreadsheetWorkbookLossManifest = {
    status: entries.length ? "lossy" : "lossless",
    entries,
  };
  if (lossManifest.status === "lossy" && options.allowLossy !== true) {
    // Reading is read-only. Carrying the counts on the loss error keeps a
    // paused confirmation from looking like a zero-cell import, while nothing
    // is written before the reader accepts the manifest.
    throw new SpreadsheetWorkbookLossError(lossManifest, {
      sheetCount: response.sheet_count,
      cellCount: response.cell_count,
      sheetNames: [...response.sheet_names],
    });
  }
  return {
    content: response.content as unknown as NativeSpreadsheetContent,
    title: response.title.trim() || file.name.replace(/\.xlsx$/i, ""),
    lossManifest,
    sheetCount: response.sheet_count,
    cellCount: response.cell_count,
    ...(response.data_source_inventory
      ? { dataSourceInventory: dataSourceInventoryFrom(response.data_source_inventory) }
      : {}),
  };
}

const CONNECTION_KINDS: ReadonlySet<string> = new Set([
  "odbc",
  "dao",
  "file",
  "web",
  "oledb",
  "text",
  "ado",
  "dsp",
]);

function dataSourceInventoryFrom(
  record: OfficeDataSourceInventoryRecord,
): SpreadsheetDataSourceInventory {
  return Object.freeze({
    connectionParts: record.connection_parts,
    queryTableParts: record.query_table_parts,
    workbookLinkParts: record.workbook_link_parts,
    linkedDataParts: record.linked_data_parts,
    workbookLinks: Object.freeze(
      record.workbook_links.map((link) =>
        Object.freeze({
          id: link.id,
          partName: link.part_name,
          displaySource: link.display_source,
          sourceKind: link.source_kind,
        }),
      ),
    ),
    queries: Object.freeze(
      (record.queries ?? []).map((query) =>
        Object.freeze({
          name: query.name,
          formula: query.formula,
          load: query.load,
          loadsToDataModel: query.loads_to_data_model,
          lastRefreshed: query.last_refreshed,
          rowCount: query.fill_count,
          errorCount: query.fill_error_count,
          connectionId: query.connection_id,
          description: query.description,
          sheetName: query.sheet_name,
          tableName: query.table_name,
          location: query.location,
          columns: Object.freeze([...query.columns]),
          sourceKind: query.source_kind,
          functions: Object.freeze([...query.functions]),
          workbookSources: Object.freeze([...query.workbook_sources]),
          queryReferences: Object.freeze([...query.query_references]),
          externalFunctions: Object.freeze([...query.external_functions]),
          culture: query.culture,
        }),
      ),
    ),
    connections: Object.freeze(
      (record.connections ?? []).map((connection) =>
        Object.freeze({
          id: connection.id,
          name: connection.name,
          description: connection.description,
          kind: (CONNECTION_KINDS.has(connection.kind)
            ? connection.kind
            : "unknown") as SpreadsheetConnectionKind,
          powerQuery: connection.power_query,
          backgroundRefresh: connection.background_refresh,
          refreshInterval: connection.refresh_interval,
          refreshOnLoad: connection.refresh_on_load,
          removeDataOnSave: connection.remove_data_on_save,
          excludeFromRefreshAll: connection.exclude_from_refresh_all,
          savePassword: connection.save_password,
          connectionFile: connection.connection_file,
          alwaysUseConnectionFile: connection.always_use_connection_file,
          connectionString: connection.connection_string,
          commandText: connection.command_text,
          commandType: connection.command_type,
          usedIn: Object.freeze(
            connection.used_in.map((use) =>
              Object.freeze({
                sheetName: use.sheet_name,
                tableName: use.table_name,
                location: use.location,
              }),
            ),
          ),
        }),
      ),
    ),
  });
}

/** Inspect only the package directory; never follows a link or runs a query. */
export async function inspectSpreadsheetDataSourceInventory(
  file: Blob,
  options: { readonly signal?: AbortSignal } = {},
): Promise<SpreadsheetDataSourceInventory> {
  const response = await requestOfficeImportProjection(file, {
    sourceKind: "xlsx",
    target: "xlsx_data_sources",
    filename: "workbook.xlsx",
    signal: options.signal,
  });
  return dataSourceInventoryFrom(response);
}

/**
 * The confirmation dialog's text. The labels live in
 * `workbook-import-losses.ts` so the server projection's own loss records
 * (`content_json.importLosses`) read from the same table.
 */
export function formatSpreadsheetWorkbookLosses(
  manifest: SpreadsheetWorkbookLossManifest,
  locale: SpreadsheetImportLossLocale,
): string {
  return describeSpreadsheetImportLosses(manifest.entries, locale).join("\n");
}
