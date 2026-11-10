/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A SmartArt member placed from its graphic's corner cell: its offsets are
 * its place in the graphic's layout and run past that cell by design
 * (`sheetSmartArtObjectsFromNativeDiagram`), because the graphic is ONE Excel
 * object. They are neither clamped to the cell, as Excel clamps a marker's
 * offset, nor dropped when the corner cell's line is deleted.
 */
export function spreadsheetObjectOffsetsLayoutRelative(object: unknown): boolean {
  if (!object || typeof object !== "object") return false;
  const record = object as { groupId?: unknown; smartArtRole?: unknown; anchor?: { kind?: unknown } };
  return (
    typeof record.groupId === "string" &&
    record.groupId.length > 0 &&
    (record.smartArtRole === "node" || record.smartArtRole === "label" || record.smartArtRole === "connector") &&
    record.anchor?.kind !== "twoCell" &&
    record.anchor?.kind !== "absolute"
  );
}
