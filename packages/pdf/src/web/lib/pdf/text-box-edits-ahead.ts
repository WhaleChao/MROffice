/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export type PdfTextBoxEditsAhead<E extends object> = Map<string, {
  /** The boxes given and not yet shown by the page, oldest first. */
  given: E[];
  /** The page's box when the oldest of them was given, or the last of them
   * the page gave back: while the page shows it, it has not moved on. */
  over: E | undefined;
}>;

/** Past this many boxes of one id the oldest go: the page would have to
 * give one of them back for it to matter, hundreds of changes late. */
const KEPT_PER_BOX = 256;

/** `shown` is `given` as the page keeps it: the same box, or a copy of it
 * with only what a write adds to a box (nothing of the box given changed). */
function pageKeepsGiven<E extends object>(given: E, shown: E): boolean {
  if (given === shown) return true;
  const kept = shown as Record<string, unknown>;
  for (const [name, value] of Object.entries(given as Record<string, unknown>)) {
    const other = kept[name];
    if (value === other) continue;
    if (value === undefined || other === undefined) return false;
    try {
      if (JSON.stringify(value) !== JSON.stringify(other)) return false;
    } catch {
      return false;
    }
  }
  return true;
}

/** Where `shown` stands among the boxes given (-1: it is the box they were
 * given over, the page has not moved), or null when it is none of them. */
function placeAmongGiven<E extends object>(
  entry: Readonly<{ given: readonly E[]; over: E | undefined }>,
  shown: E | undefined,
): number | null {
  if (shown === entry.over) return -1;
  if (shown === undefined) return null;
  const same = entry.given.lastIndexOf(shown);
  if (same >= 0) return same;
  for (let index = entry.given.length - 1; index >= 0; index -= 1) {
    if (pageKeepsGiven(entry.given[index]!, shown)) return index;
  }
  return null;
}

/** The layer gives `edit` to the page, which shows `shown` for that box. */
export function notePdfTextBoxEditGiven<E extends object>(
  ahead: PdfTextBoxEditsAhead<E>,
  id: string,
  edit: E,
  shown: E | undefined,
): void {
  const entry = ahead.get(id);
  if (!entry) {
    ahead.set(id, { given: [edit], over: shown });
    return;
  }
  entry.given.push(edit);
  if (entry.given.length > KEPT_PER_BOX) entry.given.splice(0, entry.given.length - KEPT_PER_BOX);
}

/**
 * Each box as the newest the layer knows of it, the page showing
 * `shownById`: the last box given while the page shows the box it was given
 * over or one given before it, the page's own otherwise. Reads only (a
 * render that is thrown away has settled nothing): a new map of the
 * layer's own.
 */
export function pdfTextBoxEditsWithAhead<E extends object>(
  ahead: PdfTextBoxEditsAhead<E>,
  shownById: Readonly<Record<string, E>>,
): Record<string, E> {
  const newest: Record<string, E> = { ...shownById };
  for (const [id, entry] of ahead) {
    const last = entry.given.at(-1);
    if (last && placeAmongGiven(entry, shownById[id]) !== null) newest[id] = last;
  }
  return newest;
}

/**
 * The page shows `shownById` now, for good (a render that was committed):
 * what it has given back is let go, and so is everything kept of a box it
 * shows as none of the boxes given.
 */
export function settlePdfTextBoxEditsAhead<E extends object>(
  ahead: PdfTextBoxEditsAhead<E>,
  shownById: Readonly<Record<string, E | undefined>>,
): void {
  for (const [id, entry] of ahead) {
    const shown = shownById[id];
    const at = placeAmongGiven(entry, shown);
    if (at === -1) continue;
    if (at === null) {
      ahead.delete(id);
      continue;
    }
    entry.given.splice(0, at + 1);
    entry.over = shown;
    if (entry.given.length === 0) ahead.delete(id);
  }
}
