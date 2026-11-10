/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export interface PdfNativeResourceIssue {
  readonly kind: "duplicate-font-resource" | "malformed-font-resource" | "unembedded-font" | "missing-unicode-map";
  readonly resourceName: string;
  readonly objectNumber: number;
  readonly generation: number;
}

export function readPdfNativeResourceIssue(value: unknown): PdfNativeResourceIssue | null {
  if (!value || typeof value !== "object") return null;
  const issue = value as PdfNativeResourceIssue;
  if (!["duplicate-font-resource", "malformed-font-resource", "unembedded-font", "missing-unicode-map"].includes(issue.kind) ||
    typeof issue.resourceName !== "string" || !/^[A-Za-z][A-Za-z0-9_.+\-]{0,63}$/.test(issue.resourceName) ||
    !Number.isSafeInteger(issue.objectNumber) || issue.objectNumber < 1 || issue.objectNumber > 2_147_483_647 ||
    !Number.isSafeInteger(issue.generation) || issue.generation < 0 || issue.generation > 65535) return null;
  return Object.freeze({ kind: issue.kind, resourceName: issue.resourceName, objectNumber: issue.objectNumber, generation: issue.generation });
}

const RESOURCE_ISSUE_MESSAGE_KEYS = {
  "duplicate-font-resource": "pdf.resourceIssue.duplicateFontResource",
  "malformed-font-resource": "pdf.resourceIssue.malformedFontResource",
  "unembedded-font": "pdf.resourceIssue.unembeddedFont",
  "missing-unicode-map": "pdf.resourceIssue.missingUnicodeMap",
} as const;

/**
 * The reader-facing sentence for a font resource the native editor cannot
 * use. With a translator it is the catalogue message in the reader's
 * language; without one (worker and engine errors) it stays English.
 */
export function pdfNativeResourceIssueMessage(
  issue: PdfNativeResourceIssue,
  t?: (key: string, values: Record<string, string>) => string,
): string {
  const values = {
    resource: issue.resourceName,
    object: `${issue.objectNumber}:${issue.generation}`,
  };
  if (t) return t(RESOURCE_ISSUE_MESSAGE_KEYS[issue.kind], values);
  const reasons = {
    "duplicate-font-resource": "has duplicate definitions",
    "malformed-font-resource": "has an invalid reference",
    "unembedded-font": "has no embedded font program",
    "missing-unicode-map": "has no Unicode character map",
  } as const;
  return `Font /${values.resource} (object ${values.object}) ${reasons[issue.kind]}. Only text using this font is affected.`;
}
