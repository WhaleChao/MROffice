/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export const SMARTART_MARKERLESS_LAYOUTS: ReadonlySet<string> = new Set([
  "urn:microsoft.com/office/officeart/2005/8/layout/vList2",
  "urn:microsoft.com/office/officeart/2005/8/layout/cycle1",
]);

/** Whether `block` is the words of a file's SmartArt shape in such a layout. */
export function smartArtLayoutHidesMarkers(block: unknown): boolean {
  const source = (block as { source?: { nativeSmartArtView?: { layoutIdentity?: unknown; role?: unknown } } } | null)
    ?.source;
  const view = source?.nativeSmartArtView;
  return (
    !!view &&
    view.role === "text" &&
    typeof view.layoutIdentity === "string" &&
    SMARTART_MARKERLESS_LAYOUTS.has(view.layoutIdentity)
  );
}

/**
 * The rows of `block` as Bullets / Numbering leave them on such a shape: a
 * level-1 list row keeps its list and draws no marker (`markerHidden`), as
 * PowerPoint keeps the list on the data point and its layout draws none
 * there. Any other block's rows are returned as they are.
 */
export function rowsWithLayoutHiddenMarkers<T extends { level: number; bullet: string; markerHidden?: true }>(
  block: unknown,
  rows: T[],
): T[] {
  if (!smartArtLayoutHidesMarkers(block)) return rows;
  return rows.map((row) => {
    const hidden = row.bullet !== "none" && row.level === 0;
    if (hidden === (row.markerHidden === true)) return row;
    const next = { ...row };
    if (hidden) next.markerHidden = true;
    else delete next.markerHidden;
    return next;
  });
}
