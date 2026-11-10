/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A PDF line has no paragraph indent: an indent written into it moves the
 * run's origin, and a reopened run proves its own origin as indent 0. To show
 * the indent the reader set (and to let Decrease Indent take it back), the
 * page's native edit journal is read: the records that wrote this run chain
 * by origin (each record's original state is where the previous one left the
 * run), and the indent is how far right of the FIRST record's frame the run
 * now starts. A run nobody indented has no chain and shows 0.
 */

type JournalState = { textX?: unknown; baseline?: unknown; leftIndent?: unknown; firstLineIndent?: unknown; alignment?: unknown };
export type NativeRetainedAlignment = "left" | "center" | "right";

function finite(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

function mutationOf(record: unknown): {
  frameLeft: number;
  original: { x: number; y: number };
  replacement: { x: number; y: number; indented: boolean; alignment: NativeRetainedAlignment };
} | null {
  if (!record || typeof record !== "object") return null;
  const fidelity = (record as { textFidelity?: unknown }).textFidelity;
  const mutation = fidelity && typeof fidelity === "object"
    ? (fidelity as { propertyMutation?: unknown }).propertyMutation
    // A revision published as its bytes carries it beside them.
    : (record as { kind?: unknown }).kind === "bytes-publish"
      ? (record as { propertyMutation?: unknown }).propertyMutation
      : null;
  if (!mutation || typeof mutation !== "object") return null;
  const { frame, original, replacement } = mutation as {
    frame?: { left?: unknown };
    original?: JournalState;
    replacement?: JournalState;
  };
  const frameLeft = finite(frame?.left);
  const ox = finite(original?.textX);
  const oy = finite(original?.baseline);
  const rx = finite(replacement?.textX);
  const ry = finite(replacement?.baseline);
  if (frameLeft === null || ox === null || oy === null || rx === null || ry === null) {
    return null;
  }
  const indented =
    (finite(replacement?.leftIndent) ?? 0) !== 0 ||
    (finite(replacement?.firstLineIndent) ?? 0) !== 0;
  const alignment =
    replacement?.alignment === "center" || replacement?.alignment === "right"
      ? replacement.alignment
      : "left";
  return { frameLeft, original: { x: ox, y: oy }, replacement: { x: rx, y: ry, indented, alignment } };
}

/**
 * The paragraph places a consolidation record carries (the server's rewrite
 * of a long native chain as one section keeps no records before it): each
 * run its chain left, where it stands, the frame its first edit had, the
 * alignment its newest edit wrote and whether any indented it
 * (editable_page_patchback.py `_pdf_native_paragraph_positions`).
 */
function consolidatedPositionsOf(record: unknown): Array<{
  frameLeft: number;
  at: { x: number; y: number };
  alignment: NativeRetainedAlignment;
  indented: boolean;
}> {
  if (!record || typeof record !== "object") return [];
  const kept = (record as { paragraphPositions?: unknown }).paragraphPositions;
  if ((record as { kind?: unknown }).kind !== "source-consolidation" || !Array.isArray(kept)) return [];
  return kept.flatMap((entry) => {
    const value = entry as { textX?: unknown; baseline?: unknown; frameLeft?: unknown; alignment?: unknown; indented?: unknown };
    const x = finite(value?.textX), y = finite(value?.baseline), frameLeft = finite(value?.frameLeft);
    if (x === null || y === null || frameLeft === null) return [];
    return [{
      frameLeft,
      at: { x, y },
      alignment: value.alignment === "center" || value.alignment === "right" ? value.alignment : "left",
      indented: value.indented === true,
    }];
  });
}

const NEAR = 0.002;
const same = (a: { x: number; y: number }, b: { x: number; y: number }) =>
  Math.abs(a.x - b.x) <= NEAR && Math.abs(a.y - b.y) <= NEAR;

export interface NativeRetainedParagraph {
  /** How far right of the first edit's frame the run now starts (>= 0). */
  readonly frameShift: number;
  /** The alignment the newest edit of the run wrote. */
  readonly alignment: NativeRetainedAlignment;
  /** The indent the run carries: the shift, when an edit indented it and it
   * is still left-aligned; otherwise 0. */
  readonly leftIndent: number;
}

const NONE: NativeRetainedParagraph = Object.freeze({ frameShift: 0, alignment: "left", leftIndent: 0 });

/**
 * What the journal says about this run's paragraph place: its alignment and
 * how far its frame has moved since the first edit. `origin` is the run's
 * proven text origin (textX, baseline) as reopened.
 */
export function nativeRetainedParagraph(
  journal: unknown,
  origin: { textX: number; baseline: number } | null | undefined,
): NativeRetainedParagraph {
  if (!Array.isArray(journal) || !origin) return NONE;
  const records = journal.map(mutationOf);
  let at = { x: origin.textX, y: origin.baseline };
  let first: { frameLeft: number } | null = null;
  let newest: { replacement: { alignment: NativeRetainedAlignment } } | null = null;
  let anyIndent = false;
  const used = new Set<number>();
  // Walk back from the newest record that left the run where it is now.
  for (;;) {
    let found = -1;
    for (let index = records.length - 1; index >= 0; index -= 1) {
      const record = records[index];
      if (record && !used.has(index) && same(record.replacement, at)) {
        found = index;
        break;
      }
    }
    if (found < 0) break;
    used.add(found);
    const record = records[found]!;
    newest ??= record;
    first = record;
    anyIndent ||= record.replacement.indented;
    if (same(record.original, at)) break;
    at = record.original;
  }
  // Past the first record the walk reached: a consolidation the chain was
  // rewritten into holds where its runs stood, and how they got there.
  const oldestUsed = used.size ? Math.min(...used) : journal.length;
  for (let index = oldestUsed - 1; index >= 0; index -= 1) {
    const base = consolidatedPositionsOf(journal[index]).find((entry) => same(entry.at, at));
    if (!base) continue;
    newest ??= { replacement: { alignment: base.alignment } };
    first = { frameLeft: base.frameLeft };
    anyIndent ||= base.indented;
    break;
  }
  if (!first || !newest) return NONE;
  const shift = origin.textX - first.frameLeft;
  const frameShift = shift > NEAR ? Math.round(shift * 1000) / 1000 : 0;
  const alignment = newest.replacement.alignment;
  return {
    frameShift,
    alignment,
    leftIndent: anyIndent && alignment === "left" ? frameShift : 0,
  };
}

/** The indent (points, >= 0) the journal says this run carries, or 0. */
export function nativeRetainedLeftIndent(
  journal: unknown,
  origin: { textX: number; baseline: number } | null | undefined,
): number {
  return nativeRetainedParagraph(journal, origin).leftIndent;
}

/**
 * The move that puts a run whose frame now starts `frameShift` right of where
 * it first did at alignment `target` in that first frame. The planner aligns
 * the run in its current frame [left, right], whose right edge is the first
 * frame's (the free page space); the move makes up the difference: all of
 * the shift for left, half for centre, none for right.
 */
export function nativeAlignmentMoveX(
  frameShift: number,
  target: NativeRetainedAlignment,
): number {
  if (!(frameShift > 0)) return 0;
  const move = target === "left" ? -frameShift : target === "center" ? -frameShift / 2 : 0;
  return Math.round(move * 1000) / 1000;
}
