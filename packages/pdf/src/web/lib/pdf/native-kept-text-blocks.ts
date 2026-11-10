/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { pdfNativePinnedBoxOfWrite } from "./native-page-box-map";

export interface PdfNativeKeptTextBlock {
  /** 1-based page the block is drawn on. */
  readonly pageNumber: number;
  /** The block's rows, top to bottom, as the write set them. */
  readonly rows: readonly string[];
  /** The frame a reader resized the block to, in the user space its rows
   * are drawn in (`firstLeft` is where its first row starts): what the box
   * is shown at when that space is the page's. */
  readonly frame?: Readonly<{ left: number; right: number; firstLeft: number }>;
  /** Where the block's first item stands, in PDF user space (the reader
   * item's origin): a block of one row is the box drawn from there. */
  readonly origin?: Readonly<{ x: number; y: number }>;
  /** A move set the block where it stands: words beside it are not its line. */
  readonly moved?: true;
  




  readonly pinned?: true;
}

/** How many of the journal's newest records are read. */
const KEPT_BLOCK_RECORDS = 64;

function record(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

function finite(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

function rowsOf(fidelity: Record<string, unknown>): string[] | null {
  const text = typeof fidelity.replacementDisplayText === "string" ? fidelity.replacementDisplayText : null;
  const rewrite = record(fidelity.paragraphRewrite);
  const block = record(rewrite?.block);
  if (block && text !== null && !rewrite?.history && Array.isArray(block.lines)) {
    // The rewrite's lines index the text it set, its typed breaks taken out
    // (a break is no character the block draws). Sliced from the text with
    // them, every row after a break was read a character early, matched
    // nothing on the page, and the block came apart into boxes (UN SDG p1,
    
    // edit's words read as a box of its own).
    const characters = Array.from(text.replace(/\n/gu, ""));
    const rows: string[] = [];
    for (const line of block.lines) {
      const span = record(line);
      const start = finite(span?.start);
      const end = finite(span?.end);
      if (start === null || end === null || start < 0 || end < start || end > characters.length) return null;
      rows.push(characters.slice(start, end).join("").trimEnd());
    }
    return rows;
  }
  const linePlan = record(fidelity.linePlan);
  if (linePlan && Array.isArray(linePlan.lines)) {
    const rows = new Map<number, string>();
    for (const [index, value] of linePlan.lines.entries()) {
      const line = record(value);
      if (!line || typeof line.replacementText !== "string") return null;
      const row = finite(line.lineIndex) ?? index;
      rows.set(row, (rows.get(row) ?? "") + line.replacementText);
    }
    return [...rows.entries()].sort((left, right) => left[0] - right[0]).map(([, row]) => row.trimEnd());
  }
  const positioned = record(fidelity.positionedLines);
  if (positioned && Array.isArray(positioned.lines)) {
    const rows: string[] = [];
    for (const value of positioned.lines) {
      const evidence = record(record(value)?.evidence);
      if (typeof evidence?.replacementDisplayText !== "string") return null;
      rows.push(evidence.replacementDisplayText.trimEnd());
    }
    return rows;
  }
  return null;
}

/** A reader item's origin (its transform's translation), or null. */
function itemOrigin(value: unknown): { x: number; y: number } | null {
  const transform = record(value)?.transform;
  if (!Array.isArray(transform) || transform.length < 6) return null;
  const x = finite(transform[4]);
  const y = finite(transform[5]);
  return x === null || y === null ? null : { x, y };
}

function originOf(value: unknown): Readonly<{ x: number; y: number }> | null {
  const origin = record(value);
  const x = finite(origin?.x);
  const y = finite(origin?.y);
  return x === null || y === null ? null : Object.freeze({ x, y });
}

/** The rows a selection's items draw, top to bottom: an item opens a row
 * after one the reader ended its line on, or off the row's baseline. */
function rowsOfItems(items: readonly unknown[]): string[] | null {
  const rows: string[] = [];
  let baseline: number | null = null;
  let ended = false;
  for (const value of items) {
    const item = record(value);
    const origin = itemOrigin(item);
    if (!item || !origin || typeof item.str !== "string") return null;
    const height = finite(item.height) ?? 0;
    if (baseline === null || ended || Math.abs(origin.y - baseline) > Math.max(height, 1) / 2) {
      rows.push("");
      baseline = origin.y;
    }
    rows[rows.length - 1] += item.str;
    ended = item.hasEOL === true;
  }
  return rows.map((row) => row.trimEnd());
}

/** A block a consolidation carried (the server's rewrite of a long native
 * chain keeps one record: editable_page_patchback.py
 * `_pdf_native_kept_text_blocks`), or null. */
function carriedBlock(value: unknown): PdfNativeKeptTextBlock | null {
  const block = record(value);
  const pageNumber = finite(block?.pageNumber);
  const rows = block?.rows;
  const origin = originOf(block?.origin);
  if (!block || pageNumber === null || !Number.isSafeInteger(pageNumber) || pageNumber < 1 ||
    !Array.isArray(rows) || rows.length < (origin ? 1 : 2) ||
    !rows.every((row) => typeof row === "string" && row.trim())) return null;
  const frame = record(block.frame);
  const left = finite(frame?.left);
  const right = finite(frame?.right);
  const firstLeft = finite(frame?.firstLeft);
  return Object.freeze({
    pageNumber,
    rows: Object.freeze([...(rows as string[])]),
    ...(left !== null && right !== null && firstLeft !== null && right > left
      ? { frame: Object.freeze({ left, right, firstLeft }) }
      : {}),
    ...(origin ? { origin } : {}),
    ...(block.moved === true ? { moved: true as const } : {}),
    ...(block.pinned === true && origin ? { pinned: true as const } : {}),
  });
}

/** The block one journal record set, or null: a text write's rows (a
 * re-wrapped paragraph, a line plan, positioned lines, or the one row it
 * wrote), a paragraph property edit's text, a font extension's write, a
 * move's or a group move's box where it landed. Restores are not blocks of
 * their own. */
export function pdfNativeKeptTextBlockOfRecord(entry: unknown): PdfNativeKeptTextBlock | null {
  const kept = record(entry);
  const move = record(kept?.textMove);
  const groupMove = record(kept?.textGroupMove);
  if (groupMove?.restore) return null;
  const fidelity = move
    ? record(move.sourceFidelity)
    : groupMove
      ? record(groupMove.sourceFidelity)
      // A write the font extension set (words a subset lacks, a family
      // served with them) names its paragraph in its own evidence.
      : (record(kept?.textFidelity) ?? record(record(kept?.paragraphProperties)?.textReplacement) ??
        record(record(kept?.fontExtensionText)?.sourceFidelity));
  const selection = record(fidelity?.selection);
  const pageNumber = finite(selection?.pageNumber);
  if (!fidelity || pageNumber === null || !Number.isSafeInteger(pageNumber) || pageNumber < 1) return null;
  // An undo or redo writes back what an earlier record wrote or replaced:
  // that record (or the page as it was) names the box.
  if (fidelity.positionedRestore === true ||
    ["paragraphRewrite", "positionedLines", "positionedRunGroup", "consolidatedRuns"]
      .some((key) => Boolean(record(fidelity[key])?.history))) return null;
  const items = Array.isArray(selection?.items) && selection.items.length ? selection.items : [selection?.item];
  // Where the box's first item stands after the write: a move adds its
  // vector, a group move maps it (x' = a x + e, y' = d y + f).
  let origin = itemOrigin(items[0]);
  if (origin && move) {
    const vector = originOf(move.move);
    origin = vector ? { x: origin.x + vector.x, y: origin.y + vector.y } : null;
  } else if (origin && groupMove) {
    const map = record(groupMove.transform);
    const [a, d, e, f] = [map?.a, map?.d, map?.e, map?.f].map(finite);
    origin = a !== null && d !== null && e !== null && f !== null
      ? { x: a * origin.x + e, y: d * origin.y + f }
      : null;
  }
  const written = rowsOf(fidelity);
  const display = typeof fidelity.replacementDisplayText === "string" ? fidelity.replacementDisplayText : null;
  const ownRows = written ??
    (groupMove ? rowsOfItems(items) : display !== null && !display.includes("\n") ? [display.trimEnd()] : null);
  const rows = ownRows && !move && !groupMove ? withFontExtensionText(kept, fidelity, ownRows) : ownRows;
  if (!rows || rows.length === 0 || rows.some((row) => !row.trim())) return null;
  // A block of one row, or one a move set, is named by where it stands too:
  // one row of text alone may be drawn more than once on a page.
  const located = rows.length < 2 || move !== null || groupMove !== null;
  if (located && !origin) return null;
  const block = record(record(fidelity.paragraphRewrite)?.block);
  const left = finite(block?.frameLeft);
  const right = finite(block?.frameRight);
  const firstLeft = finite(block?.firstLeft);
  const upright = Array.isArray(block?.basis) && block.basis.join(",") === "1,0,0,1";
  return Object.freeze({
    pageNumber,
    rows: Object.freeze(rows),
    ...(!move && !groupMove && block?.frameRequested === true && upright && left !== null && right !== null &&
      firstLeft !== null && right > left
      ? { frame: Object.freeze({ left, right, firstLeft }) }
      : {}),
    ...(located && origin ? { origin: Object.freeze(origin) } : {}),
    ...(move !== null || groupMove !== null ? { moved: true as const } : {}),
  });
}









function withFontExtensionText(kept: Record<string, unknown> | null, fidelity: Record<string, unknown>,
  rows: readonly string[]): string[] | null {
  const extension = record(kept?.fontExtensionText);
  if (!extension || record(fidelity.paragraphRewrite) || extension.sourceFidelity !== fidelity) return [...rows];
  const after = typeof extension.replacementText === "string" ? extension.replacementText : null;
  const before = typeof fidelity.replacementText === "string" ? fidelity.replacementText : null;
  if (after === null || before === null || after === before) return [...rows];
  if (after.includes("\n") || Array.isArray(extension.replacementLines)) return null;
  const holding = rows.flatMap((row, index) => row.includes(before.trimEnd()) ? [index] : []);
  if (holding.length !== 1 || rows[holding[0]].split(before.trimEnd()).length !== 2) return null;
  return rows.map((row, index) => index === holding[0] ? row.replace(before.trimEnd(), after.trimEnd()) : row);
}

/**
 * The box a text write set, named by what it wrote from where its first item
 * stood (native-page-box-map.ts pdfNativePinnedBoxOfWrite): kept with a write
 * whose record names no rows of its own (a line retyped in place, a run's
 * properties), so the box reads back whole however its new text wraps (a
 * title shortened in place came back as two boxes, 1005.1585 p1). Null for a
 * move, a restore, an undo or redo, and a write that names no selection.
 */
export function pdfNativeWrittenBoxBlockOfRecord(entry: unknown): PdfNativeKeptTextBlock | null {
  const kept = record(entry);
  if (!kept || kept.textMove || kept.textGroupMove) return null;
  const extension = record(kept.fontExtensionText);
  const fidelity = record(kept.textFidelity) ?? record(record(kept.paragraphProperties)?.textReplacement) ??
    record(extension?.sourceFidelity);
  const selection = record(fidelity?.selection);
  const pageNumber = finite(selection?.pageNumber);
  if (!fidelity || pageNumber === null || !Number.isSafeInteger(pageNumber) || pageNumber < 1) return null;
  if (fidelity.positionedRestore === true ||
    ["paragraphRewrite", "positionedLines", "positionedRunGroup", "consolidatedRuns"]
      .some((key) => Boolean(record(fidelity[key])?.history))) return null;
  const items = Array.isArray(selection?.items) && selection.items.length ? selection.items : [selection?.item];
  if (!items.every((item) => record(item) && typeof record(item)!.str === "string" && itemOrigin(item))) return null;
  // Only a write of the whole box names it: a word of a line set apart, named
  // as a box of its own, would be taken out of the line it stands in.
  const shown = record(selection?.item)?.str;
  if (typeof shown !== "string" || finite(selection?.startOffset) !== 0 || finite(selection?.endOffset) !== shown.length) {
    return null;
  }
  const text = typeof extension?.replacementText === "string" && extension.sourceFidelity === fidelity
    ? extension.replacementText
    : typeof fidelity.replacementDisplayText === "string" ? fidelity.replacementDisplayText
      : typeof fidelity.replacementText === "string" ? fidelity.replacementText : null;
  return text === null ? null
    : pdfNativePinnedBoxOfWrite(pageNumber, items as unknown as Parameters<typeof pdfNativePinnedBoxOfWrite>[1], text);
}

/** The most blocks a published revision carries (its box and its page's
 * map, native-page-box-map.ts). */
export const PDF_NATIVE_PUBLISHED_KEPT_BLOCKS_MAX = 1_000;

/** The journal's text blocks, newest first. A revision published as its
 * bytes carries the blocks its edit set (`keptTextBlocks`: the box it wrote
 * first, then its page's box map): read from every such record, however old,
 * as a page's map stays the boxes of the page while the page draws them. */
export function pdfNativeKeptTextBlocks(journal: unknown): readonly PdfNativeKeptTextBlock[] {
  if (!Array.isArray(journal)) return [];
  const blocks: PdfNativeKeptTextBlock[] = [];
  const newest = journal.length - KEPT_BLOCK_RECORDS;
  for (let index = journal.length - 1; index >= 0; index -= 1) {
    const kept = record(journal[index]);
    if (kept?.kind === "source-consolidation") {
      // Carried oldest first: newest first here, after the records since.
      const carried = Array.isArray(kept.keptTextBlocks) ? kept.keptTextBlocks : [];
      for (const value of [...carried].reverse()) {
        const block = carriedBlock(value);
        if (block) blocks.push(block);
      }
      continue;
    }
    if (kept?.kind === "bytes-publish") {
      const carried = Array.isArray(kept.keptTextBlocks) ? kept.keptTextBlocks : [];
      for (const value of carried.slice(0, PDF_NATIVE_PUBLISHED_KEPT_BLOCKS_MAX)) {
        const block = carriedBlock(value);
        if (block) blocks.push(block);
      }
      continue;
    }
    if (index < newest) continue;
    const block = pdfNativeKeptTextBlockOfRecord(kept);
    if (block) blocks.push(block);
  }
  return Object.freeze(blocks);
}
