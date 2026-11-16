/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * FILE-LEVEL PDF operations: MERGE several separate PDF files into one, and
 * SPLIT one PDF into multiple files by page-range groups.
 *
 * This is distinct from `lib/pdf/page-ops.ts`, which manages WITHIN-document
 * structural edit STATE (rotate / delete / reorder / insert-blank) and produces
 * render props — it never combines or carves separate documents. THIS module is
 * the missing Acrobat-parity capability: it actually writes new PDF byte
 * streams with pdf-lib (`PDFDocument.create()` + `copyPages` from each source),
 * mirroring the only-place-we-touch-pdf-lib convention of `flatten-apply.ts`.
 *
 * All page math (which 0-based indices each output carries) is delegated to the
 * pure, tested `lib/pdf/range-spec.ts` so this module stays a thin pdf-lib
 * adapter: load → copyPages → addPage → save.
 */

import { PDFDocument } from "pdf-lib";

import { loadPdfForMutation } from "./encrypted-guard";

import {
  parseRangeGroups,
  parseRangeSpec,
  rangeGroupToZeroBased,
  splitEveryN,
  type RangeGroup,
} from "./range-spec";
import { pdfLibObjectsPerTick } from "@mrnote/core/web/lib/pdf/pdf-lib-pacing.ts";

/** One step in an insert plan: pull a page from a given source by 0-based index. */
export interface InsertPlanStep {
  /** Which document the page comes from. */
  source: "base" | "insert";
  /** 0-based page index within that source. */
  index: number;
}

/**
 * PURE assembly of an "insert pages from another PDF at a position" plan.
 *
 * Produces the ORDERED list of (source, index) steps that build the result:
 * the first `position` pages of the base, then the selected `insertIndices`
 * from the inserted doc (in the given order, after de-duping + bounds-checking),
 * then the remaining base pages. `position` is a 1-based INSERTION POINT (1 =
 * before page 1, `baseCount + 1` = after the last page) and is clamped to
 * `[1, baseCount + 1]` so an out-of-range value lands at the nearest edge
 * (Acrobat's "Insert → Before/After" behaviour). `insertIndices` are 0-based
 * indices into the inserted document; out-of-range / duplicate / non-integer
 * indices are dropped (a duplicate keeps its FIRST occurrence so the caller can
 * pass a free-form selection). When no insert index survives, the plan is just
 * the base pages unchanged (the caller can treat that as "nothing to insert").
 *
 * No pdf-lib / DOM: the page math is fully unit-testable in isolation.
 */
export function insertionPlan(
  baseCount: number,
  insertCount: number,
  insertIndices: ReadonlyArray<number>,
  position: number,
): InsertPlanStep[] {
  const base = Number.isFinite(baseCount) && baseCount > 0 ? Math.trunc(baseCount) : 0;
  const insTotal =
    Number.isFinite(insertCount) && insertCount > 0 ? Math.trunc(insertCount) : 0;
  // Clamp the 1-based insertion point into [1, base + 1]; default to the end
  // (after the last page) when it is non-finite.
  const rawPos = Number(position);
  let pos = Number.isFinite(rawPos) ? Math.trunc(rawPos) : base + 1;
  if (pos < 1) pos = 1;
  if (pos > base + 1) pos = base + 1;
  // De-dupe + bounds-check the inserted indices, preserving first-seen order.
  const seen = new Set<number>();
  const inserts: number[] = [];
  for (const raw of insertIndices) {
    if (!Number.isInteger(raw)) continue;
    if (raw < 0 || raw >= insTotal) continue;
    if (seen.has(raw)) continue;
    seen.add(raw);
    inserts.push(raw);
  }
  const steps: InsertPlanStep[] = [];
  // Base pages BEFORE the insertion point (0-based 0 .. pos-2).
  for (let i = 0; i < pos - 1; i += 1) steps.push({ source: "base", index: i });
  // The inserted pages, in selection order.
  for (const idx of inserts) steps.push({ source: "insert", index: idx });
  // Base pages AT / AFTER the insertion point (0-based pos-1 .. base-1).
  for (let i = pos - 1; i < base; i += 1) steps.push({ source: "base", index: i });
  return steps;
}

/** A named source PDF for merge (name only flows into the merged filename). */
export interface MergeSource {
  /** Display name (e.g. the uploaded file name); used only for diagnostics. */
  name: string;
  /** The raw PDF bytes. */
  bytes: Uint8Array | ArrayBuffer;
}

/** One produced output file from a split. */
export interface SplitOutput {
  /** The new PDF bytes for this range group. */
  bytes: Uint8Array;
  /** The originating range group (label / pages / index) for filename + UI. */
  group: RangeGroup;
}

export interface MergeOptions {
  /**
   * Skip sources that fail to load instead of aborting the whole merge. When a
   * source is skipped, its name is collected in {@link MergeResult.skipped}.
   * Defaults to true so one corrupt file never blocks combining the rest.
   */
  skipUnreadable?: boolean;
}

export interface MergeResult {
  /** The merged PDF bytes. */
  bytes: Uint8Array;
  /** Names of sources that could not be loaded and were skipped. */
  skipped: string[];
  /** Total pages copied into the merged document. */
  pageCount: number;
}

/** Copy raw bytes into a fresh ArrayBuffer-backed view (Blob-safe, detached-safe). */
export function toBlobBytes(bytes: Uint8Array): Uint8Array {
  const copy = new Uint8Array(bytes.byteLength);
  copy.set(bytes);
  return copy;
}

/**
 * Merge several PDFs (in the given ORDER) into a single document: create a fresh
 * doc, then for each source copy ALL of its pages and append them. Returns the
 * merged bytes plus the names of any sources that were skipped because they
 * could not be parsed.
 *
 * At least one source must load successfully, otherwise this throws — an empty
 * (zero-page) PDF is never produced.
 */
export async function mergePdfs(
  sources: ReadonlyArray<MergeSource>,
  options: MergeOptions = {},
): Promise<MergeResult> {
  const skipUnreadable = options.skipUnreadable !== false;
  const out = await PDFDocument.create();
  const skipped: string[] = [];
  let loadedAny = false;

  for (const source of sources) {
    let src: PDFDocument;
    try {
      src = await loadPdfForMutation(source.bytes);
    } catch (err) {
      if (skipUnreadable) {
        skipped.push(source.name);
        continue;
      }
      throw new Error(
        `Could not read PDF "${source.name}": ${
          err instanceof Error ? err.message : String(err)
        }`,
      );
    }
    const indices = src.getPageIndices(); // all pages, in order
    const copied = await out.copyPages(src, indices);
    for (const page of copied) out.addPage(page);
    loadedAny = true;
  }

  if (!loadedAny) {
    throw new Error("No readable PDF sources to merge");
  }

  const bytes = await out.save({ objectsPerTick: pdfLibObjectsPerTick() });
  return { bytes: toBlobBytes(bytes), skipped, pageCount: out.getPageCount() };
}

/** The result of inserting pages from another PDF into this one. */
export interface InsertResult {
  /** The new PDF bytes: base pages with the inserted pages spliced in. */
  bytes: Uint8Array;
  /** Total pages in the produced document. */
  pageCount: number;
  /** How many pages were actually inserted (after subset + bounds filtering). */
  insertedCount: number;
}

/**
 * INSERT pages FROM another PDF into THIS one at a chosen position, optionally
 * limited to a page SUBSET of the inserted file (Acrobat "Insert Pages" / WPS
 * 插入页面). Unlike {@link mergePdfs} (which only appends whole files at the
 * end), this splices a selection at an arbitrary 1-based `position`:
 *
 *   - `position` is the 1-based insertion point in the BASE doc: 1 inserts
 *     before page 1, `baseCount + 1` (or any larger value) appends at the end.
 *     Clamped to a valid edge — never throws on an out-of-range value.
 *   - `insertSpec` is an OPTIONAL Acrobat-style 1-based range spec
 *     (`"1-3, 5, 8-"`) selecting which pages of the inserted file to bring in,
 *     in that order. Omitted / blank → ALL pages of the inserted file.
 *
 * The page math is the pure, tested {@link insertionPlan}; this function is the
 * thin pdf-lib adapter that copies each planned page from the right source.
 * Throws only when a document can't be parsed (a corrupt file is a hard error
 * here — there is nothing to fall back to, unlike a multi-source merge).
 */
export async function insertPagesFromSource(
  baseBytes: Uint8Array | ArrayBuffer,
  insertBytes: Uint8Array | ArrayBuffer,
  position: number,
  insertSpec?: string,
): Promise<InsertResult> {
  const base = await loadPdfForMutation(baseBytes);
  const insert = await loadPdfForMutation(insertBytes);
  const baseCount = base.getPageCount();
  const insertCount = insert.getPageCount();

  // Resolve the inserted page SUBSET: a non-blank spec selects 1-based pages
  // (parsed/clamped/sorted by the shared range parser) → 0-based indices; a
  // blank/absent spec means the whole inserted file in natural order.
  const spec = typeof insertSpec === "string" ? insertSpec.trim() : "";
  const insertIndices =
    spec === ""
      ? insert.getPageIndices()
      : parseRangeSpec(spec, insertCount).map((p) => p - 1);

  const plan = insertionPlan(baseCount, insertCount, insertIndices, position);

  const out = await PDFDocument.create();
  // Copy each source's needed pages ONCE (pdf-lib copyPages de-dups internally
  // per call), then add them in the planned order. We copy per distinct source
  // index so a page reused (it can't be, plan de-dups) or reordered still maps.
  const baseIdx = plan.filter((s) => s.source === "base").map((s) => s.index);
  const insIdx = plan.filter((s) => s.source === "insert").map((s) => s.index);
  const copiedBase = baseIdx.length ? await out.copyPages(base, baseIdx) : [];
  const copiedInsert = insIdx.length ? await out.copyPages(insert, insIdx) : [];
  // Map each (source,index) back to its copied page via a running cursor per
  // source, since copyPages returns pages in the SAME order as the index list.
  const baseByIndex = new Map<number, (typeof copiedBase)[number]>();
  baseIdx.forEach((idx, i) => baseByIndex.set(idx, copiedBase[i]));
  const insertByIndex = new Map<number, (typeof copiedInsert)[number]>();
  insIdx.forEach((idx, i) => insertByIndex.set(idx, copiedInsert[i]));
  for (const step of plan) {
    const page =
      step.source === "base"
        ? baseByIndex.get(step.index)
        : insertByIndex.get(step.index);
    if (page) out.addPage(page);
  }

  const bytes = await out.save({ objectsPerTick: pdfLibObjectsPerTick() });
  return {
    bytes: toBlobBytes(bytes),
    pageCount: out.getPageCount(),
    insertedCount: insIdx.length,
  };
}

/**
 * Build ONE new PDF from an explicit list of 0-based source page indices, IN THE
 * GIVEN ORDER. Indices out of bounds for the source are silently skipped (a
 * final guard against a stale page count). Returns null when no index is
 * in-bounds (so a caller never produces an empty zero-page PDF). The single
 * place page-copy assembly happens — both group-per-file split and combined
 * extract route through it.
 */
export async function buildDocFromIndices(
  source: PDFDocument,
  indices: ReadonlyArray<number>,
): Promise<Uint8Array | null> {
  const total = source.getPageCount();
  const inBounds = indices.filter((i) => Number.isInteger(i) && i >= 0 && i < total);
  if (inBounds.length === 0) return null;
  const out = await PDFDocument.create();
  const copied = await out.copyPages(source, inBounds);
  for (const page of copied) out.addPage(page);
  const bytes = await out.save({ objectsPerTick: pdfLibObjectsPerTick() });
  return toBlobBytes(bytes);
}

/**
 * Build ONE new PDF for a single range group from already-loaded source pages.
 * Thin wrapper over {@link buildDocFromIndices} using the group's 0-based
 * indices. Returns null when the group resolves to zero in-bounds pages.
 */
async function buildGroupDoc(
  source: PDFDocument,
  group: RangeGroup,
): Promise<Uint8Array | null> {
  return buildDocFromIndices(source, rangeGroupToZeroBased(group));
}

const DEFAULT_SPLIT_AGGREGATE_BYTE_BUDGET = 256 * 1024 * 1024;

export interface SplitPdfOptions {
  /**
   * Maximum bytes retained across every generated split output. The producer
   * enforces this after each group, before generating the next one.
   */
  readonly maxAggregateBytes?: number;
}

function readSplitAggregateByteBudget(options?: SplitPdfOptions): number {
  const budget =
    options?.maxAggregateBytes ?? DEFAULT_SPLIT_AGGREGATE_BYTE_BUDGET;
  if (!Number.isSafeInteger(budget) || budget < 1) {
    throw new Error("Invalid split PDF aggregate byte budget");
  }
  return budget;
}

async function buildSplitOutputs(
  source: PDFDocument,
  groups: ReadonlyArray<RangeGroup>,
  options?: SplitPdfOptions,
): Promise<SplitOutput[]> {
  const budget = readSplitAggregateByteBudget(options);
  const outputs: SplitOutput[] = [];
  let aggregateBytes = 0;
  try {
    for (const group of groups) {
      const bytes = await buildGroupDoc(source, group);
      if (!bytes) continue;
      aggregateBytes += bytes.byteLength;
      if (!Number.isSafeInteger(aggregateBytes) || aggregateBytes > budget) {
        bytes.fill(0);
        throw new Error("Split PDF aggregate byte budget exceeded");
      }
      outputs.push({ bytes, group });
    }
    return outputs;
  } catch (error) {
    for (const output of outputs) output.bytes.fill(0);
    outputs.length = 0;
    throw error;
  }
}

/**
 * Split one PDF into multiple output files, one per supplied {@link RangeGroup}.
 * Use {@link splitBySpec} / {@link splitByEveryN} to build the groups from user
 * input. Groups that resolve to no in-bounds pages are dropped. Each output
 * carries its originating group so the caller can name the file.
 */
export async function splitPdf(
  sourceBytes: Uint8Array | ArrayBuffer,
  groups: ReadonlyArray<RangeGroup>,
  options?: SplitPdfOptions,
): Promise<SplitOutput[]> {
  const source = await loadPdfForMutation(sourceBytes);
  return buildSplitOutputs(source, groups, options);
}

/**
 * Split one PDF by an Acrobat-style range spec (`"1-3, 5, 8-"`), producing one
 * output file PER comma-token range. Convenience wrapper over
 * {@link parseRangeGroups} + {@link splitPdf}. Returns `[]` when the spec
 * resolves to nothing.
 */
export async function splitBySpec(
  sourceBytes: Uint8Array | ArrayBuffer,
  spec: string,
): Promise<SplitOutput[]> {
  const source = await loadPdfForMutation(sourceBytes);
  const groups = parseRangeGroups(spec, source.getPageCount());
  return buildSplitOutputs(source, groups);
}

/** The combined-extract result: one new PDF plus the page list it carries. */
export interface ExtractResult {
  /** The new PDF bytes containing exactly the selected pages, in page order. */
  bytes: Uint8Array;
  /** The selected 1-based source page numbers, sorted + de-duped. */
  pages: number[];
}

/**
 * EXTRACT the pages an Acrobat-style spec (`"1,3,5"` / `"1-3, 8-"`) selects into
 * ONE combined PDF — the union of every comma-token, sorted + de-duped, written
 * as a single document (NOT one file per range). This is the "combine selected
 * pages into one file" affordance: `1,3,5` yields a single 3-page PDF rather than
 * three 1-page files (which is {@link splitBySpec}).
 *
 * Uses the pure, tested {@link parseRangeSpec} for the page math (it flattens all
 * tokens into one ascending list) and {@link buildDocFromIndices} for assembly.
 * Returns null when the spec resolves to no in-bounds pages (empty / all-garbage
 * input), so the caller can surface "selects no pages" rather than emit an empty
 * download.
 */
export async function extractBySpec(
  sourceBytes: Uint8Array | ArrayBuffer,
  spec: string,
): Promise<ExtractResult | null> {
  const source = await loadPdfForMutation(sourceBytes);
  const pages = parseRangeSpec(spec, source.getPageCount());
  if (pages.length === 0) return null;
  const indices = pages.map((p) => p - 1);
  const bytes = await buildDocFromIndices(source, indices);
  if (!bytes) return null;
  return { bytes, pages };
}

/**
 * Split one PDF into fixed-size chunks of `chunkSize` pages each ("split every
 * N pages"). Convenience wrapper over {@link splitEveryN} + {@link splitPdf}.
 */
export async function splitByEveryN(
  sourceBytes: Uint8Array | ArrayBuffer,
  chunkSize: number,
): Promise<SplitOutput[]> {
  const source = await loadPdfForMutation(sourceBytes);
  const groups = splitEveryN(source.getPageCount(), chunkSize);
  return buildSplitOutputs(source, groups);
}
