/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * `boxes` with those in `drawn` (the ids drawn last, in the order drawn)
 * standing in that order, in the places `boxes` gives to them; every other
 * box where it was given. `boxes` itself when it has them so already.
 */
export function pdfTextBoxesInDrawnOrder<T extends Readonly<{ id: string }>>(
  boxes: T[],
  drawn: readonly string[],
): T[] {
  if (boxes.length < 2 || drawn.length < 2) return boxes;
  const place = new Map<string, number>();
  drawn.forEach((id, index) => place.set(id, index));
  let last = -1;
  let ordered = true;
  for (const box of boxes) {
    const at = place.get(box.id);
    if (at === undefined) continue;
    if (at < last) {
      ordered = false;
      break;
    }
    last = at;
  }
  if (ordered) return boxes;
  const kept = boxes
    .filter((box) => place.has(box.id))
    .sort((a, b) => (place.get(a.id) ?? 0) - (place.get(b.id) ?? 0));
  let next = 0;
  return boxes.map((box) => (place.has(box.id) ? kept[next++]! : box));
}

/** Whether `ids` is the order `boxes` stand in. */
export function pdfTextBoxDrawnOrderIs(
  boxes: ReadonlyArray<Readonly<{ id: string }>>,
  ids: readonly string[],
): boolean {
  if (boxes.length !== ids.length) return false;
  for (let index = 0; index < boxes.length; index += 1) {
    if (boxes[index]!.id !== ids[index]) return false;
  }
  return true;
}
