/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export const MAX_DOCUMENT_PARAGRAPH_SPACING_PT = 1584;

/**
 * Word's Design > Paragraph Spacing gallery (Document Formatting group), in
 * gallery order. Word desktop applies each built-in set as the document's
 * paragraph defaults: space before, space after (points) and an "auto" line
 * multiple. The values are Word's own built-in table (twips/240ths in the
 * Word for Mac binary: 0/0/240, 0/80/240, 0/120/276, 0/200/276, 0/120/360,
 * 0/160/480), which the gallery's tooltip reads back as "Before / After /
 * Line Spacing". `default` clears the document's override so the template
 * (and MRAI's own prose defaults) apply again, like Word's "Default" row.
 */
export const DOCUMENT_PARAGRAPH_SPACING_PRESETS = Object.freeze([
  { id: "default", beforePt: null, afterPt: null, lineHeight: null },
  { id: "none", beforePt: 0, afterPt: 0, lineHeight: "1" },
  { id: "compact", beforePt: 0, afterPt: 4, lineHeight: "1" },
  { id: "tight", beforePt: 0, afterPt: 6, lineHeight: "1.15" },
  { id: "open", beforePt: 0, afterPt: 10, lineHeight: "1.15" },
  { id: "relaxed", beforePt: 0, afterPt: 6, lineHeight: "1.5" },
  { id: "double", beforePt: 0, afterPt: 8, lineHeight: "2" },
] as const);

export type DocumentParagraphSpacingPreset =
  (typeof DOCUMENT_PARAGRAPH_SPACING_PRESETS)[number];

export type DocumentParagraphSpacingPresetId =
  DocumentParagraphSpacingPreset["id"];

export function normalizeDocumentParagraphSpacing(
  value: unknown,
): number | null {
  const raw =
    typeof value === "number" ? value : Number.parseFloat(String(value ?? ""));
  if (!Number.isFinite(raw)) return null;
  return Math.min(
    MAX_DOCUMENT_PARAGRAPH_SPACING_PT,
    Math.max(0, Math.round(raw)),
  );
}

function normalizePresetLineHeight(value: unknown): string | null {
  if (value === null || value === undefined || value === "") return null;
  const numeric =
    typeof value === "number" ? value : Number.parseFloat(String(value));
  return Number.isFinite(numeric) ? String(numeric) : null;
}

export function documentParagraphSpacingPreset(
  id: unknown,
): DocumentParagraphSpacingPreset | null {
  return (
    DOCUMENT_PARAGRAPH_SPACING_PRESETS.find((preset) => preset.id === id) ??
    null
  );
}

/**
 * The gallery row the document's current defaults match exactly, or null
 * (Word shows no checked row for a custom combination). All three values take
 * part: Tight and Relaxed share their after spacing and differ only in line
 * spacing.
 */
export function documentParagraphSpacingPresetFor(
  beforePt: unknown,
  afterPt: unknown,
  lineHeight: unknown,
): DocumentParagraphSpacingPresetId | null {
  const before = normalizeDocumentParagraphSpacing(beforePt);
  const after = normalizeDocumentParagraphSpacing(afterPt);
  const line = normalizePresetLineHeight(lineHeight);
  return (
    DOCUMENT_PARAGRAPH_SPACING_PRESETS.find(
      (preset) =>
        preset.beforePt === before &&
        preset.afterPt === after &&
        normalizePresetLineHeight(preset.lineHeight) === line,
    )?.id ?? null
  );
}
