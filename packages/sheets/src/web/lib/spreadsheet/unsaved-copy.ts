/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * "Download a copy" of a workbook the window could not save (Excel's Save a
 * Copy while an upload is failing or blocked): the window's own content is
 * sent to the server's unsaved-copy export, which converts it to .xlsx and
 * stores nothing (read access to the page is all it needs). When that is
 * not possible (signed out, offline, refused) the person still gets their
 * data: every sheet as a CSV file, zipped, built here.
 */
import { zipSync, strToU8 } from "fflate";

import { apiFetchRaw } from "@mrnote/core/web/lib/api.ts";
import { downloadBlobFile, startDownloadWhenAllowed } from "@mrnote/core/web/lib/download-file.ts";
import { toDelimitedText } from "@mrnote/core/web/lib/spreadsheet/csv.ts";

export interface UnsavedCopyResult {
  kind: "xlsx" | "csv-zip";
  /** What the .xlsx could not carry (the server's X-MRAI-Export-Losses). */
  losses: string[];
  /** How many items that is in all (X-MRAI-Export-Loss-Count). */
  lossCount: number;
  /** Why the .xlsx was not available (csv-zip only). */
  fallbackReason?: string;
}

export function safeWorkbookFileName(title: string, fallback: string): string {
  return title.trim().replace(/[\\/:*?"<>|]+/g, "-").replace(/\.+$/g, "").slice(0, 120) || fallback;
}

function readLosses(header: string | null): string[] {
  if (!header) return [];
  try {
    const parsed = JSON.parse(header) as unknown;
    if (Array.isArray(parsed)) return parsed.map(String).filter(Boolean);
  } catch {
    // a plain list
  }
  return header.split(",").map((entry) => entry.trim()).filter(Boolean);
}

interface CsvSheet {
  name?: unknown;
  cells?: unknown;
}

function cellText(cell: unknown): string {
  if (cell === null || cell === undefined) return "";
  if (typeof cell !== "object") return String(cell);
  const record = cell as { display?: unknown; value?: unknown; formula?: unknown };
  if (typeof record.display === "string") return record.display;
  if (record.value !== undefined && record.value !== null) return String(record.value);
  return typeof record.formula === "string" ? record.formula : "";
}

/** Every sheet of the content as CSV text, by sheet name (unique). */
export function workbookContentToCsvFiles(content: Record<string, unknown>): Array<{ name: string; csv: string }> {
  const sheets = Array.isArray(content.sheets) ? (content.sheets as CsvSheet[]) : [];
  const used = new Set<string>();
  return sheets.map((sheet, index) => {
    const cells = sheet.cells && typeof sheet.cells === "object" ? (sheet.cells as Record<string, unknown>) : {};
    let maxRow = -1;
    let maxCol = -1;
    const grid = new Map<string, string>();
    for (const [key, cell] of Object.entries(cells)) {
      const [row, col] = key.split(":").map(Number);
      if (!Number.isInteger(row) || !Number.isInteger(col) || row < 0 || col < 0) continue;
      const text = cellText(cell);
      if (!text) continue;
      grid.set(`${row}:${col}`, text);
      maxRow = Math.max(maxRow, row);
      maxCol = Math.max(maxCol, col);
    }
    const rows: string[][] = [];
    for (let row = 0; row <= maxRow; row += 1) {
      const line: string[] = [];
      for (let col = 0; col <= maxCol; col += 1) line.push(grid.get(`${row}:${col}`) ?? "");
      rows.push(line);
    }
    const base = safeWorkbookFileName(typeof sheet.name === "string" ? sheet.name : "", `Sheet${index + 1}`);
    let name = base;
    for (let suffix = 2; used.has(name.toLowerCase()); suffix += 1) name = `${base} (${suffix})`;
    used.add(name.toLowerCase());
    return { name, csv: `﻿${toDelimitedText(rows, ",", true)}` };
  });
}

export async function downloadUnsavedWorkbookCopy(input: {
  pageId: string | null;
  content: Record<string, unknown>;
  fileName: string;
}): Promise<UnsavedCopyResult> {
  let fallbackReason = "no-page";
  if (input.pageId) {
    try {
      const response = await apiFetchRaw(
        `/api/v1/pages/${encodeURIComponent(input.pageId)}/exports/unsaved-copy`,
        { method: "POST", body: JSON.stringify({ content_json: input.content, format: "xlsx" }) },
        { contentType: "application/json", suppressUnauthorizedHandling: true },
      );
      const blob = await response.blob();
      const name = `${input.fileName}.xlsx`;
      if (blob.size > 0 && startDownloadWhenAllowed(name, () => downloadBlobFile(name, blob))) {
        const losses = readLosses(response.headers.get("X-MRAI-Export-Losses"));
        const counted = Number(response.headers.get("X-MRAI-Export-Loss-Count"));
        return { kind: "xlsx", losses, lossCount: Number.isSafeInteger(counted) && counted > 0 ? counted : losses.length };
      }
      fallbackReason = "empty";
    } catch (error) {
      fallbackReason = error instanceof Error ? error.message : "failed";
    }
  }
  const files = workbookContentToCsvFiles(input.content);
  const entries: Record<string, Uint8Array> = {};
  for (const file of files.length ? files : [{ name: "Sheet1", csv: "﻿" }]) {
    entries[`${file.name}.csv`] = strToU8(file.csv);
  }
  const zipped = zipSync(entries);
  const blob = new Blob([zipped.buffer.slice(zipped.byteOffset, zipped.byteOffset + zipped.byteLength) as ArrayBuffer], { type: "application/zip" });
  const zipName = `${input.fileName} (CSV).zip`;
  if (!startDownloadWhenAllowed(zipName, () => downloadBlobFile(zipName, blob))) throw new Error("download_unavailable");
  return { kind: "csv-zip", losses: [], lossCount: 0, fallbackReason };
}
