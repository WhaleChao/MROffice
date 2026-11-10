/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The notice Excel shows before it opens a .csv or .txt file that one of the
 * enabled Automatic Data Conversion options would change (Microsoft 365
 * Insider, "Control data conversions in Excel for Windows and Mac": "Excel
 * displays a warning message when it detects that at least one of the
 * optional automatic data conversions is enabled and about to occur when
 * opening a .csv or .txt file ... the ability to open the file once without
 * converting the data"). It lists the conversions it found; Convert opens the
 * file as the options say, Don't Convert opens it once with none of the four,
 * and its tick box turns the notice off (the fifth option).
 *
 * The page import runs outside React, so the notice is asked through a
 * registered host (the workspace mounts one); with no host, or with the
 * notice switched off, the file opens as the options say, as Excel's does.
 */
import type { DataConversionKind } from "@mrnote/core/web/lib/spreadsheet/data-conversion-options.ts";

export interface CsvConversionNoticeRequest {
  readonly fileName: string;
  /** What opening the file would convert, in Excel's list order. */
  readonly conversions: readonly DataConversionKind[];
}

export interface CsvConversionNoticeAnswer {
  readonly convert: boolean;
  /** "Don't notify me about automatic data conversions" was ticked. */
  readonly stopNotifying: boolean;
}

export type CsvConversionNoticeAsker = (
  request: CsvConversionNoticeRequest,
) => Promise<CsvConversionNoticeAnswer>;

let asker: CsvConversionNoticeAsker | null = null;

/** The workspace's notice host; returns its unregistration. */
export function registerCsvConversionNotice(next: CsvConversionNoticeAsker): () => void {
  asker = next;
  return () => {
    if (asker === next) asker = null;
  };
}

export function csvConversionNoticeAsker(): CsvConversionNoticeAsker | null {
  return asker;
}

/** The event a spreadsheet window sends to open File > Options > Data's
 *  Automatic Data Conversion box (the workspace host draws it). */
export const OPEN_DATA_CONVERSION_OPTIONS_EVENT = "mrai:open-data-conversion-options";

export function openDataConversionOptions(): void {
  if (typeof window !== "undefined") {
    window.dispatchEvent(new CustomEvent(OPEN_DATA_CONVERSION_OPTIONS_EVENT));
  }
}
