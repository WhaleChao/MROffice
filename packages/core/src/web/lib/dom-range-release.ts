/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Taking a fresh Range a measurement has read out of its document's range
 * registry (moved here from paged/measurementRange so the text engine can use
 * it): a Range stays registered with its document until it is garbage
 * collected, and every node removal in that document visits every registered
 * range to move its boundary points.
 */

/** A document of this window's own, never shown and never changed: where a
 * fresh range is parked once it is read (`releaseFreshRange`). */
const parkingDocuments = new WeakMap<Document, Document | null>();

function parkingDocumentFor(doc: Document): Document | null {
  if (parkingDocuments.has(doc)) return parkingDocuments.get(doc) ?? null;
  let parking: Document | null = null;
  try {
    parking = doc.implementation.createHTMLDocument("");
  } catch {
    parking = null;
  }
  parkingDocuments.set(doc, parking);
  return parking;
}

/**
 * A fresh range a measurement has read and will not read again, taken out of
 * its document's registry: its boundary is set in a private document, which
 * moves the range there (a range lives in the document of its boundary
 * points). What it measured is unchanged; the document's later node removals
 * no longer visit it until it is collected (see above: a fresh range per
 * character, kept on purpose where a measurement's order matters, left
 * thousands registered, and every page move emptied and filled sheets past
 * all of them).
 */
export function releaseFreshRange(range: Range): void {
  const parking = parkingDocumentFor(range.startContainer.ownerDocument ?? (range.startContainer as Document));
  if (!parking) return;
  try {
    range.setStart(parking, 0);
    range.collapse(true);
  } catch {
    // A range that cannot move stays where it is, as before.
  }
}

/** `window.__docsFreshRangesKept = true`: the ranges the text engine's quick
 * line check and a table cell's line boxes read are left registered, as
 * before (A/B; the punctuation hang's ranges are released either way). */
export function freshRangesKept(): boolean {
  return (globalThis as { __docsFreshRangesKept?: boolean }).__docsFreshRangesKept === true;
}
