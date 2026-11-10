/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Client for the workbook change log (Review > Show Changes).
 *
 * GET  /api/v1/pages/{id}/spreadsheet-changes  newest first, optionally narrowed
 *      to a sheet or a range, keyset-paged with `before`.
 * POST /api/v1/pages/{id}/spreadsheet-changes  this client's committed edits;
 *      the server stamps the author and the time.
 */

import { apiGet, apiPost } from "./api";
import type {
  SpreadsheetChangeActor,
  SpreadsheetChangeCell,
  SpreadsheetChangeEntry,
  SpreadsheetChangeFilter,
  SpreadsheetChangeKind,
  SpreadsheetChangeRecord,
} from "./spreadsheet/change-log";

export interface SpreadsheetChangesPage {
  items: SpreadsheetChangeEntry[];
  nextCursor: string | null;
  retentionDays: number;
}

const KINDS = new Set<SpreadsheetChangeKind>([
  "edit",
  "undo",
  "redo",
  "insert_rows",
  "delete_rows",
  "insert_columns",
  "delete_columns",
]);

function record(value: unknown): Record<string, unknown> | null {
  return value !== null && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

function wholeNumber(value: unknown): number | null {
  return typeof value === "number" && Number.isInteger(value) && value >= 0
    ? value
    : null;
}

function text(value: unknown): string | null {
  return typeof value === "string" ? value : null;
}

function actorFromWire(value: unknown): SpreadsheetChangeActor {
  const wire = record(value) ?? {};
  return {
    userId: text(wire["user_id"]),
    name: text(wire["display_name"]),
    avatarKind: text(wire["avatar_kind"]),
    avatarDoodle: text(wire["avatar_doodle"]),
    avatarUrlSmall: text(wire["avatar_url_small"]),
    accentColor: text(wire["accent_color"]),
  };
}

function cellFromWire(value: unknown): SpreadsheetChangeCell | null {
  const wire = record(value);
  if (!wire) return null;
  const r = wholeNumber(wire["r"]);
  const c = wholeNumber(wire["c"]);
  const before = text(wire["before"]);
  const after = text(wire["after"]);
  if (r === null || c === null || before === null || after === null) return null;
  return { r, c, before, after };
}

/**
 * A change's time in epoch milliseconds, keeping the microseconds the server
 * writes: one upload stamps its edits a microsecond apart, in the order they
 * were made, and Date.parse alone rounds them onto one millisecond, so the
 * pane fell back on the (random) row ids and listed an edit, its Undo and its
 * Redo out of order once saved (live audit S-093).
 */
export function changeTimeFromWire(value: string): number {
  const millis = Date.parse(value);
  if (!Number.isFinite(millis)) return millis;
  const fraction = /T[^.]*\.(\d+)/.exec(value)?.[1] ?? "";
  const beyondMillis = fraction.slice(3, 6);
  return beyondMillis ? millis + Number(beyondMillis.padEnd(3, "0")) / 1000 : millis;
}

/** One server row, or null when it does not read as one. */
export function spreadsheetChangeFromWire(value: unknown): SpreadsheetChangeEntry | null {
  const wire = record(value);
  if (!wire) return null;
  const id = text(wire["id"]);
  const clientChangeId = text(wire["client_change_id"]);
  const sheetId = text(wire["sheet_id"]);
  const kind = text(wire["kind"]) as SpreadsheetChangeKind | null;
  const r1 = wholeNumber(wire["r1"]);
  const c1 = wholeNumber(wire["c1"]);
  const r2 = wholeNumber(wire["r2"]);
  const c2 = wholeNumber(wire["c2"]);
  const cellCount = wholeNumber(wire["cell_count"]);
  const createdAt = changeTimeFromWire(text(wire["created_at"]) ?? "");
  if (
    !id ||
    !clientChangeId ||
    !sheetId ||
    !kind ||
    !KINDS.has(kind) ||
    r1 === null ||
    c1 === null ||
    r2 === null ||
    c2 === null ||
    cellCount === null ||
    !Number.isFinite(createdAt) ||
    !Array.isArray(wire["cells"])
  ) {
    return null;
  }
  const cells = wire["cells"]
    .map(cellFromWire)
    .filter((cell): cell is SpreadsheetChangeCell => cell !== null);
  return {
    id,
    clientChangeId,
    sheetId,
    sheetName: text(wire["sheet_name"]) ?? "",
    kind,
    r1,
    c1,
    r2,
    c2,
    cellCount,
    cells,
    truncated: wire["truncated"] === true,
    createdAt,
    actor: actorFromWire(wire["actor"]),
  };
}

function itemsFromWire(value: unknown): SpreadsheetChangeEntry[] {
  const items = record(value)?.["items"];
  if (!Array.isArray(items)) throw new Error("spreadsheet_changes_response_invalid");
  return items
    .map(spreadsheetChangeFromWire)
    .filter((entry): entry is SpreadsheetChangeEntry => entry !== null);
}

export async function listSpreadsheetChanges(
  pageId: string,
  query: { filter: SpreadsheetChangeFilter; before?: string | null; limit?: number },
  signal?: AbortSignal,
): Promise<SpreadsheetChangesPage> {
  const params = new URLSearchParams();
  const { filter } = query;
  if (filter.scope !== "workbook") params.set("sheet_id", filter.sheetId);
  if (filter.scope === "range") {
    params.set("r1", String(filter.r1));
    params.set("c1", String(filter.c1));
    params.set("r2", String(filter.r2));
    params.set("c2", String(filter.c2));
  }
  if (query.before) params.set("before", query.before);
  if (query.limit) params.set("limit", String(query.limit));
  const search = params.toString();
  const payload = await apiGet<unknown>(
    `/api/v1/pages/${encodeURIComponent(pageId)}/spreadsheet-changes${search ? `?${search}` : ""}`,
    signal ? { signal } : undefined,
  );
  const wire = record(payload);
  return {
    items: itemsFromWire(payload),
    nextCursor: text(wire?.["next_cursor"]) || null,
    retentionDays: wholeNumber(wire?.["retention_days"]) ?? 365,
  };
}

/** Upload committed edits; returns them as the server stored them. */
export async function recordSpreadsheetChanges(
  pageId: string,
  records: readonly SpreadsheetChangeRecord[],
  signal?: AbortSignal,
): Promise<SpreadsheetChangeEntry[]> {
  const payload = await apiPost<unknown>(
    `/api/v1/pages/${encodeURIComponent(pageId)}/spreadsheet-changes`,
    {
      changes: records.map((change) => ({
        client_change_id: change.clientChangeId,
        sheet_id: change.sheetId,
        sheet_name: change.sheetName.slice(0, 255),
        kind: change.kind,
        r1: change.r1,
        c1: change.c1,
        r2: change.r2,
        c2: change.c2,
        cell_count: change.cellCount,
        cells: change.cells,
        truncated: change.truncated,
      })),
    },
    signal ? { signal } : undefined,
  );
  return itemsFromWire(payload);
}
