/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { normalizeRotation, type PdfRotation } from "@mrnote/core/web/lib/pdf/page-ops.ts";

export type PdfPageOrganizerOperation =
  | {
      op: "rotate_pdf_page";
      page: number;
      degrees: PdfRotation;
    }
  | {
      op: "delete_pdf_page";
      page: number;
    }
  | {
      op: "move_pdf_page";
      page: number;
      after_page: number;
    }
  // The page grid shows the inserted (blank) pages too: their changes are
  // the page model's own operations for them, and an order that moves one
  // is written whole (inserted pages by id).
  | {
      op: "reorder_pdf_pages";
      page_count: number;
      sequence: Array<{ page: number } | { inserted_id: string }>;
    }
  | { op: "delete_pdf_inserted_page"; id: string }
  | { op: "rotate_pdf_inserted_page"; id: string; degrees: 90 | 180 | 270 };

export interface PdfPageOrganizerDraft {
  /** Visible original pages in their proposed output order. */
  order: number[];
  /** Proposed absolute /Rotate value for every visible original page. */
  rotations: Record<number, PdfRotation>;
  /** Original pages newly removed in this organizer session. */
  deletedPages: number[];
}

function uniquePages(pages: ReadonlyArray<number>): number[] {
  const seen = new Set<number>();
  const result: number[] = [];
  for (const page of pages) {
    if (!Number.isInteger(page) || page < 1 || seen.has(page)) continue;
    seen.add(page);
    result.push(page);
  }
  return result;
}

export function createPdfPageOrganizerDraft(options: {
  order: ReadonlyArray<number>;
  rotations?: Readonly<Record<number, number>>;
}): PdfPageOrganizerDraft {
  const order = uniquePages(options.order);
  const rotations: Record<number, PdfRotation> = {};
  for (const page of order) {
    rotations[page] = normalizeRotation(options.rotations?.[page]);
  }
  return { order, rotations, deletedPages: [] };
}

export function rotatePdfPageOrganizerPages(
  draft: PdfPageOrganizerDraft,
  pages: ReadonlyArray<number>,
  quarterTurns: number,
): PdfPageOrganizerDraft {
  const turns = Number.isFinite(quarterTurns) ? Math.trunc(quarterTurns) : 0;
  if (((turns % 4) + 4) % 4 === 0) return draft;
  const visible = new Set(draft.order);
  const rotations = { ...draft.rotations };
  let changed = false;
  for (const page of uniquePages(pages)) {
    if (!visible.has(page)) continue;
    const next = normalizeRotation(
      normalizeRotation(rotations[page]) + turns * 90,
    );
    if (next === rotations[page]) continue;
    rotations[page] = next;
    changed = true;
  }
  return changed ? { ...draft, rotations } : draft;
}

export interface PdfPageOrganizerDeleteResult {
  draft: PdfPageOrganizerDraft;
  rejected: boolean;
}

/**
 * Delete selected original pages while preserving the PDF invariant that at
 * least one original page remains. The refusal is explicit so the UI can show
 * a visible explanation instead of silently ignoring the command.
 */
export function deletePdfPageOrganizerPages(
  draft: PdfPageOrganizerDraft,
  pages: ReadonlyArray<number>,
): PdfPageOrganizerDeleteResult {
  const requested = new Set(uniquePages(pages));
  const deleted = draft.order.filter((page) => requested.has(page));
  if (deleted.length === 0) return { draft, rejected: false };
  if (draft.order.length - deleted.length < 1) {
    return { draft, rejected: true };
  }
  const deletedPages = uniquePages([...draft.deletedPages, ...deleted]).sort(
    (a, b) => a - b,
  );
  return {
    draft: {
      order: draft.order.filter((page) => !requested.has(page)),
      rotations: { ...draft.rotations },
      deletedPages,
    },
    rejected: false,
  };
}

/**
 * Move every selected page one slot earlier/later. Discontiguous selections
 * keep their relative order, matching multi-selection movement in desktop PDF
 * organizers.
 */
export function movePdfPageOrganizerPages(
  draft: PdfPageOrganizerDraft,
  pages: ReadonlyArray<number>,
  direction: -1 | 1,
): PdfPageOrganizerDraft {
  const selected = new Set(uniquePages(pages));
  if (selected.size === 0) return draft;
  const order = [...draft.order];
  let changed = false;
  if (direction < 0) {
    for (let index = 1; index < order.length; index += 1) {
      if (!selected.has(order[index]) || selected.has(order[index - 1])) {
        continue;
      }
      [order[index - 1], order[index]] = [order[index], order[index - 1]];
      changed = true;
    }
  } else {
    for (let index = order.length - 2; index >= 0; index -= 1) {
      if (!selected.has(order[index]) || selected.has(order[index + 1])) {
        continue;
      }
      [order[index], order[index + 1]] = [order[index + 1], order[index]];
      changed = true;
    }
  }
  return changed ? { ...draft, order } : draft;
}

/**
 * Drag a selected page group before/after a visible target. Selected pages keep
 * their document order; dropping onto a member of the moving group is a no-op.
 */
export function reorderPdfPageOrganizerPages(
  draft: PdfPageOrganizerDraft,
  pages: ReadonlyArray<number>,
  targetPage: number,
  dropAfter: boolean,
): PdfPageOrganizerDraft {
  const selected = new Set(uniquePages(pages));
  if (
    selected.size === 0 ||
    selected.has(targetPage) ||
    !draft.order.includes(targetPage)
  ) {
    return draft;
  }
  const moving = draft.order.filter((page) => selected.has(page));
  if (moving.length === 0) return draft;
  const remaining = draft.order.filter((page) => !selected.has(page));
  const targetIndex = remaining.indexOf(targetPage);
  if (targetIndex < 0) return draft;
  const insertAt = targetIndex + (dropAfter ? 1 : 0);
  const order = [
    ...remaining.slice(0, insertAt),
    ...moving,
    ...remaining.slice(insertAt),
  ];
  return order.every((page, index) => page === draft.order[index])
    ? draft
    : { ...draft, order };
}

/**
 * Convert a complete organizer draft into the smallest deterministic native
 * command batch needed to reach it from the currently-visible original order.
 * Deletes are emitted first, rotations second, then only the moves required to
 * transform the surviving order. This makes the full Apply gesture one atomic
 * undo/writeback operation without replacing unrelated PDF objects.
 */
export function buildPdfPageOrganizerOperations(options: {
  initialOrder: ReadonlyArray<number>;
  initialRotations?: Readonly<Record<number, number>>;
  draft: PdfPageOrganizerDraft;
}): PdfPageOrganizerOperation[] {
  const initialOrder = uniquePages(options.initialOrder);
  const deletedSet = new Set(uniquePages(options.draft.deletedPages));
  const targetOrder = uniquePages(options.draft.order).filter(
    (page) => !deletedSet.has(page),
  );
  const targetSet = new Set(targetOrder);
  const operations: PdfPageOrganizerOperation[] = [];

  for (const page of initialOrder) {
    if (deletedSet.has(page)) {
      operations.push({ op: "delete_pdf_page", page });
    }
  }

  const initialRotations = options.initialRotations ?? {};
  for (const page of targetOrder) {
    const before = normalizeRotation(initialRotations[page]);
    const after = normalizeRotation(options.draft.rotations[page]);
    if (before !== after) {
      operations.push({ op: "rotate_pdf_page", page, degrees: after });
    }
  }

  const working = initialOrder.filter((page) => targetSet.has(page));
  for (
    let targetIndex = 0;
    targetIndex < targetOrder.length;
    targetIndex += 1
  ) {
    const page = targetOrder[targetIndex];
    if (working[targetIndex] === page) continue;
    const currentIndex = working.indexOf(page);
    if (currentIndex < 0) continue;
    working.splice(currentIndex, 1);
    working.splice(targetIndex, 0, page);
    operations.push({
      op: "move_pdf_page",
      page,
      after_page: targetIndex === 0 ? 0 : targetOrder[targetIndex - 1],
    });
  }

  return operations;
}

/** The page grid numbers an inserted (blank) page from here: a stand-in for
 * its id, never a page of the file (the grid's model is page numbers). */
export const PDF_PAGE_ORGANIZER_INSERTED_BASE = 1_000_000;

/**
 * The grid's operations, with its inserted pages' stand-in numbers turned
 * into the page model's own operations for them: a rotation or deletion of
 * one is its inserted-page operation, and an order with inserted pages that
 * anything moved in is written whole (reorder_pdf_pages: each inserted page
 * after the page before it on screen, as the grid shows it -- moved one by
 * one, an inserted page went with the original it was anchored after).
 */
export function pdfPageOrganizerOperationsWithInserted(options: {
  operations: ReadonlyArray<PdfPageOrganizerOperation>;
  /** The grid's order when the operations were made, stand-ins included. */
  initialOrder: ReadonlyArray<number>;
  /** Stand-in number to inserted page id. */
  insertedIds: ReadonlyMap<number, string>;
  /** The document's original page count (reorder_pdf_pages). */
  pageCount: number;
}): PdfPageOrganizerOperation[] {
  const { insertedIds } = options;
  if (insertedIds.size === 0) return [...options.operations];
  const out: PdfPageOrganizerOperation[] = [];
  let moved = false;
  let order = [...options.initialOrder];
  for (const operation of options.operations) {
    if (operation.op === "delete_pdf_page") {
      const id = insertedIds.get(operation.page);
      out.push(id ? { op: "delete_pdf_inserted_page", id } : operation);
      order = order.filter((page) => page !== operation.page);
    } else if (operation.op === "rotate_pdf_page") {
      const id = insertedIds.get(operation.page);
      if (!id) {
        out.push(operation);
        continue;
      }
      // A stand-in opens upright: the grid's absolute turn is the turn.
      const degrees = normalizeRotation(operation.degrees);
      if (degrees !== 0) out.push({ op: "rotate_pdf_inserted_page", id, degrees: degrees as 90 | 180 | 270 });
    } else if (operation.op === "move_pdf_page") {
      moved = true;
      const at = order.indexOf(operation.page);
      if (at < 0) continue;
      order.splice(at, 1);
      const after = operation.after_page === 0 ? -1 : order.indexOf(operation.after_page);
      order.splice(after + 1, 0, operation.page);
    } else {
      out.push(operation);
    }
  }
  if (moved) {
    out.push({
      op: "reorder_pdf_pages",
      page_count: options.pageCount,
      sequence: order.map((page) => {
        const id = insertedIds.get(page);
        return id ? { inserted_id: id } : { page };
      }),
    });
  }
  return out;
}

/** WPS's 页面选择 (CT_Page PageSelect): which pages a preset picks. */
export type PdfPageSelectionPreset =
  | "all"
  | "odd"
  | "even"
  | "landscape"
  | "portrait";

export const PDF_PAGE_SELECTION_PRESETS: readonly PdfPageSelectionPreset[] = [
  "all",
  "odd",
  "even",
  "landscape",
  "portrait",
];

/**
 * The pages of `order` a preset picks. Odd and even count the pages as they
 * now stand (the first page shown is page 1); landscape and portrait read
 * the page as it shows, its rotation applied. A square page is neither.
 */
export function selectPdfPagesByPreset(options: {
  order: ReadonlyArray<number>;
  sizes: ReadonlyMap<number, { width: number; height: number }>;
  rotations: Readonly<Record<number, PdfRotation>>;
  preset: PdfPageSelectionPreset;
}): number[] {
  const { order, sizes, rotations, preset } = options;
  return order.filter((page, index) => {
    if (preset === "all") return true;
    if (preset === "odd") return index % 2 === 0;
    if (preset === "even") return index % 2 === 1;
    const size = sizes.get(page);
    if (!size) return false;
    const quarter = normalizeRotation(rotations[page]) % 180 !== 0;
    const width = quarter ? size.height : size.width;
    const height = quarter ? size.width : size.height;
    return preset === "landscape" ? width > height : height > width;
  });
}
