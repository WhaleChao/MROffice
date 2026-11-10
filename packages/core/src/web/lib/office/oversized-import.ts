/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Read-only preview of an Office file too large to be an editable page.
 *
 * The server imports such a file as a page that shows the part of it that
 * fits — each sheet's complete top rows, a deck's leading slides, a
 * document's leading blocks — and marks it with a top-level `oversizedImport`
 * entry (apps/api/app/services/resources/oversized_import.py). The entry's
 * presence locks the page: the server refuses every content write, and the
 * original file remains the page's download and export.
 *
 * This module only reads the entry for display. Presence alone is what locks a
 * page, so a malformed entry still reads as a preview, just without detail.
 */

export const OVERSIZED_IMPORT_KEY = "oversizedImport";

const MARKER_SCHEMA = "mrnote.oversized-import.v1";
// The most sheets a page holds (the page envelope's maxSheetsTotal).
const MAX_SHEETS = 4096;
const MAX_SHEET_ID_CHARS = 200;

export type OversizedImportSourceKind = "xlsx" | "pptx" | "docx" | "file";

/** The page export that hands back the file this page was imported from. */
export const OVERSIZED_IMPORT_EXPORT_FORMAT: Record<
  Exclude<OversizedImportSourceKind, "file">,
  string
> = {
  xlsx: "xlsx",
  pptx: "pptx",
  docx: "docx",
};

/** The stored upload a preview (or an original-only page) was made from. */
export interface OversizedImportOriginal {
  readonly dataItemId: string;
  readonly filename: string;
  readonly sizeBytes: number;
  readonly mediaType: string;
}

/** Why an original-only page shows none of its file. */
export type OversizedImportOriginalReason = `encrypted:${string}` | "too_large" | "not_convertible" | "unreadable" | "workspace_full";

export interface OversizedImportSheetCut {
  readonly sheetId: string;
  /** Rows 1..rowsShown are complete; nothing below them is on the page. */
  readonly rowsShown: number;
  readonly rowCount: number;
}

export interface OversizedImportPreview {
  /** Which kind of file this page previews, when the marker says so. */
  readonly sourceKind: OversizedImportSourceKind | null;
  /** Workbooks: only the sheets that were cut; every other sheet is complete. */
  readonly sheets: readonly OversizedImportSheetCut[];
  /** Workbooks: formula cells show the values saved in the file, not live results. */
  readonly formulasShowSavedValues: boolean;
  /** Decks: leading slides shown. Documents: leading blocks shown. */
  readonly shown: number | null;
  /** How many the file itself has. */
  readonly total: number | null;
  /** The stored upload, downloaded as it is; absent on older previews,
   * which download through the page's export. */
  readonly original: OversizedImportOriginal | null;
  /** Original-only pages: why nothing of the file is shown. */
  readonly reason: OversizedImportOriginalReason | null;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isCount(value: unknown): value is number {
  return typeof value === "number" && Number.isSafeInteger(value) && value >= 0;
}

const EMPTY_PREVIEW: OversizedImportPreview = Object.freeze({
  sourceKind: null,
  sheets: Object.freeze([]),
  formulasShowSavedValues: false,
  shown: null,
  total: null,
  original: null,
  reason: null,
});

function readOriginal(value: unknown): OversizedImportOriginal | null {
  if (!isRecord(value)) return null;
  const { dataItemId, filename, sizeBytes, mediaType } = value;
  if (
    typeof dataItemId !== "string" ||
    !dataItemId ||
    dataItemId.length > 64 ||
    typeof filename !== "string" ||
    !filename ||
    !isCount(sizeBytes) ||
    typeof mediaType !== "string"
  ) {
    return null;
  }
  return Object.freeze({ dataItemId, filename, sizeBytes, mediaType });
}

/** Whether a page's content is a locked read-only preview. */
export function isOversizedImportContent(content: unknown): boolean {
  return isRecord(content) && Object.hasOwn(content, OVERSIZED_IMPORT_KEY);
}

function readSheetCuts(
  entries: readonly unknown[],
  wholeAllowed: boolean,
): OversizedImportSheetCut[] | null {
  const sheets: OversizedImportSheetCut[] = [];
  const seen = new Set<string>();
  for (const entry of entries) {
    if (!isRecord(entry)) return null;
    const { sheetId, rowsShown, rowCount } = entry;
    if (
      typeof sheetId !== "string" ||
      sheetId.length === 0 ||
      sheetId.length > MAX_SHEET_ID_CHARS ||
      seen.has(sheetId) ||
      !isCount(rowsShown) ||
      !isCount(rowCount) ||
      rowsShown > rowCount ||
      (rowsShown === rowCount && !wholeAllowed)
    ) {
      return null;
    }
    seen.add(sheetId);
    sheets.push(Object.freeze({ sheetId, rowsShown, rowCount }));
  }
  return sheets;
}

function readPartCounts(
  marker: Record<string, unknown>,
  shownKey: string,
  countKey: string,
  wholeAllowed: boolean,
): { shown: number; total: number } | null {
  const shown = marker[shownKey];
  const total = marker[countKey];
  // A preview that names its original was made from a file past its editable
  // size (or with parts too large to load left out), so it may show all.
  if (
    !isCount(shown) ||
    !isCount(total) ||
    shown < 1 ||
    shown > total ||
    (shown === total && !wholeAllowed)
  ) {
    return null;
  }
  return { shown, total };
}

/** The preview description of a page's content, or null for an ordinary page. */
export function readOversizedImportPreview(
  content: unknown,
): OversizedImportPreview | null {
  if (!isOversizedImportContent(content)) return null;
  const marker = (content as Record<string, unknown>)[OVERSIZED_IMPORT_KEY];
  if (!isRecord(marker) || marker.schema !== MARKER_SCHEMA) return EMPTY_PREVIEW;
  const original = readOriginal(marker.original);
  if (marker.sourceKind === "file") {
    const reason =
      marker.reason === "too_large" ||
      marker.reason === "not_convertible" ||
      marker.reason === "unreadable" ||
      marker.reason === "workspace_full" ||
      (typeof marker.reason === "string" && /^encrypted:[a-z0-9_]{1,64}$/.test(marker.reason))
        ? (marker.reason as OversizedImportOriginalReason)
        : null;
    return Object.freeze({ ...EMPTY_PREVIEW, sourceKind: "file", original, reason });
  }
  if (marker.sourceKind === "pptx" || marker.sourceKind === "docx") {
    const counts =
      marker.sourceKind === "pptx"
        ? readPartCounts(marker, "slidesShown", "slideCount", original !== null)
        : readPartCounts(marker, "blocksShown", "blockCount", original !== null);
    if (counts === null) return Object.freeze({ ...EMPTY_PREVIEW, original });
    return Object.freeze({
      ...EMPTY_PREVIEW,
      sourceKind: marker.sourceKind,
      shown: counts.shown,
      total: counts.total,
      original,
    });
  }
  if (
    marker.sourceKind !== "xlsx" ||
    !Array.isArray(marker.sheets) ||
    marker.sheets.length > MAX_SHEETS
  ) {
    return EMPTY_PREVIEW;
  }
  const sheets = readSheetCuts(marker.sheets, original !== null);
  if (sheets === null) return Object.freeze({ ...EMPTY_PREVIEW, original });
  return Object.freeze({
    ...EMPTY_PREVIEW,
    sourceKind: "xlsx",
    sheets: Object.freeze(sheets),
    formulasShowSavedValues: marker.formulas === "cached-values",
    original,
  });
}
