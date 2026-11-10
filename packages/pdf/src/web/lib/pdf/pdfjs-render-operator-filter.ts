/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
type Operators = Readonly<{ fnArray: ArrayLike<number>; argsArray: ArrayLike<unknown> }>;
type Ops = Readonly<Record<string, number | undefined>>;
/** What pdf.js's render() answers; its internal task holds the list it runs. */
type RenderTaskLike = { readonly _internalRenderTask?: { readonly operatorList?: Operators | null } | null };

export interface PdfRenderOperationsFilter {
  /** The filter to pass as render's `operationsFilter`. */
  readonly filter: (renderIndex: number) => boolean;
  /** The task render() answered, bound before it runs its first chunk. */
  bind(task: unknown): void;
  /** After the render: true when every rendered operation was matched to
   * the read list and the read list was used up exactly. */
  finish(): boolean;
}

function lengthOf(value: unknown): number {
  return value && typeof (value as ArrayLike<unknown>).length === "number" ? (value as ArrayLike<unknown>).length : -1;
}

export function createPdfRenderOperationsFilter(input: {
  /** The list getOperatorList answered (same intent and annotation mode as the render). */
  list: Operators;
  ops: Ops;
  /** The filter in read-list indices. */
  keep: (index: number) => boolean;
}): PdfRenderOperationsFilter {
  const { list, ops, keep } = input;
  const read = list.fnArray;
  const readArgs = list.argsArray;
  let task: RenderTaskLike | null = null;
  let failed = false;
  let i = 0; // next read operation to account for
  const decisions: boolean[] = [];
  const at = (index: number) => (index < read.length ? read[index] : -1);
  const fail = () => { failed = true; return true; };
  /** `count` groups of save, transform, `paint`, restore from `i`, all kept
   * or all left out (a folded group is drawn or not as one). */
  const groups = (count: number, paint: number | undefined): boolean | null => {
    if (!(count > 0) || typeof paint !== "number") return null;
    let decision: boolean | null = null;
    for (let group = 0; group < count; group += 1) {
      const base = i + group * 4;
      if (at(base) !== ops.save || at(base + 1) !== ops.transform || at(base + 2) !== paint || at(base + 3) !== ops.restore) return null;
      const kept = keep(base + 2);
      if (decision === null) decision = kept;
      else if (decision !== kept) return null;
    }
    i += count * 4;
    return decision;
  };
  const alignNext = (rendered: Operators, j: number): boolean => {
    const fn = rendered.fnArray[j];
    const args = rendered.argsArray[j] as unknown[] | null | undefined;
    // A one-show text object folded into the run before it: its endText,
    // beginText and setFont are gone; its setTextMatrix and show remain.
    while (fn === ops.setTextMatrix && at(i) === ops.endText && at(i + 1) === ops.beginText && at(i + 2) === ops.setFont &&
      at(i + 3) === ops.setTextMatrix) i += 3;
    if (fn === ops.constructPath && at(i) === ops.save && at(i + 1) === ops.transform && at(i + 2) === ops.constructPath &&
      at(i + 3) === ops.restore) {
      decisions[j] = keep(i + 2);
      i += 4;
      return true;
    }
    let folded: boolean | null | undefined;
    if (fn === ops.paintImageXObjectRepeat) folded = groups(lengthOf(args?.[3]) / 2, ops.paintImageXObject);
    else if (fn === ops.paintImageMaskXObjectRepeat) folded = groups(lengthOf(args?.[5]) / 2, ops.paintImageMaskXObject);
    else if (fn === ops.paintImageMaskXObjectGroup) folded = groups(lengthOf(args?.[0]), ops.paintImageMaskXObject);
    else if (fn === ops.paintInlineImageXObjectGroup) folded = groups(lengthOf(args?.[1]), ops.paintInlineImageXObject);
    if (folded !== undefined) {
      if (folded === null) return false;
      decisions[j] = folded;
      return true;
    }
    if (i >= read.length || fn !== read[i]) return false;
    decisions[j] = keep(i);
    i += 1;
    return true;
  };
  const rendered = (): Operators | null => task?._internalRenderTask?.operatorList ?? null;
  // A render that shows no list of its own (a test double of pdf.js, or a
  // build without the private field) is filtered in read-list indices, as
  // before this helper: the two lists are then taken to be the same.
  let legacy = false;
  const filter = (renderIndex: number): boolean => {
    if (failed) return true;
    const list = rendered();
    if (!list && decisions.length === 0) { legacy = true; return keep(renderIndex); }
    if (legacy) return keep(renderIndex);
    if (!list || renderIndex >= list.fnArray.length) return fail();
    for (let j = decisions.length; j <= renderIndex; j += 1) {
      if (!alignNext(list, j)) return fail();
    }
    return decisions[renderIndex] ?? fail();
  };
  return {
    filter,
    bind(value: unknown) { task = (value as RenderTaskLike) ?? null; },
    finish() {
      if (failed) return false;
      if (legacy) return true;
      const list = rendered();
      // No list of its own and never consulted (a double that renders
      // nothing): read-list indices stand, as above.
      if (!list) return decisions.length === 0;
      for (let j = decisions.length; j < list.fnArray.length; j += 1) {
        if (!alignNext(list, j)) return false;
      }
      return i === read.length && decisions.length === list.fnArray.length;
    },
  };
}
