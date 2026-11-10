/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export type PdfTextPageSeparator = "blank-line" | "form-feed";

export interface PdfTextOutputOptions {
  readonly pageSeparator: PdfTextPageSeparator;
  readonly includePageLabels: boolean;
  readonly utf8Bom: boolean;
}

export interface PdfTextOutputPage {
  readonly ordinal: number;
  readonly text: string;
}

export const DEFAULT_PDF_TEXT_OUTPUT_OPTIONS: PdfTextOutputOptions =
  Object.freeze({
    pageSeparator: "blank-line",
    includePageLabels: false,
    utf8Bom: false,
  });

function pageLabel(
  ordinal: number,
  label: string | ((ordinal: number) => string),
): string {
  if (typeof label === "function") return label(ordinal);
  return label.toLowerCase().startsWith("zh")
    ? `第 ${ordinal} 页`
    : `Page ${ordinal}`;
}

/**
 * Format the exact UTF-8 TXT payload shown in the review dialog and sent to
 * the browser download. A terminal newline is deterministic; the optional BOM
 * is the sole byte-order marker and never appears for an empty selection.
 */
export function formatPdfExtractedText(
  pages: readonly PdfTextOutputPage[],
  options: PdfTextOutputOptions,
  /** The page label's words: the UI catalogue's `(ordinal) => "Page n"`, or a
   * locale for the built-in zh/en pair. */
  locale: string | ((ordinal: number) => string),
): string {
  if (pages.length === 0) return "";
  const separator = options.pageSeparator === "form-feed" ? "\n\f\n" : "\n\n";
  const body = pages
    .map((page) =>
      options.includePageLabels
        ? `${pageLabel(page.ordinal, locale)}\n${page.text}`
        : page.text,
    )
    .join(separator);
  if (!body.trim()) return "";
  const terminated = body.endsWith("\n") ? body : `${body}\n`;
  return options.utf8Bom ? `\uFEFF${terminated}` : terminated;
}
