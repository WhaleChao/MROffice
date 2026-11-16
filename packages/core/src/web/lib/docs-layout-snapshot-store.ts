"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The layouts kept on the server for a long document
 * (app/routers/document_layout_snapshots.py): read as the page opens, written
 * once its pagination has settled. Every operation is best-effort: no layout
 * kept, or none read in time, and the page is laid out as it always was.
 */
import { apiGet, apiPut } from "./api";
import type { ProjectionNodeRecipe } from "../components/console/editor/native-document/paged/paginationProjectionReceipt";

export type LayoutEngine = "chromium" | "webkit";

export interface LayoutSnapshotRow {
  readonly layout: string;
  readonly engine: LayoutEngine;
  readonly fontsKey: string;
  readonly part: number;
  /** A part kept in segments (app/services/document_layout_snapshots): this
   * row's place, their count, the part's block it starts at, and the capture
   * all of them come from. Absent for a part kept whole. */
  readonly segment?: number;
  readonly segments?: number;
  readonly firstBlock?: number;
  readonly captureId?: string;
  readonly contentHash: string;
  readonly blockCount: number;
  readonly continues: boolean;
  readonly pages: ProjectionNodeRecipe[][];
  readonly measures: ((number[] | null)[])[];
  /** Per block of each page, the engine's line corrections its textblocks
   * drew (engineLinesKept); absent when none were noted. */
  readonly lines?: ((number[] | null)[])[];
  /** Per block of each page, per textblock, Auto line spacing's boxes it drew
   * ([digest of its model, boxes]; engineLinesKept). */
  readonly autoLines?: (unknown[] | null)[][];
  /** On a part put together from its segments (layoutSnapshot's
   * wholePartRows): each segment's first block, block count and content. */
  readonly segmentHashes?: readonly (readonly [first: number, count: number, hash: string])[];
}

/** The layout code of this build (next.config.mjs), "" outside a build. */
export const LAYOUT_CODE_VERSION: string = process.env.NEXT_PUBLIC_DOCS_LAYOUT_VERSION || "";

export function layoutEngineName(): LayoutEngine {
  if (typeof navigator === "undefined") return "chromium";
  const agent = navigator.userAgent;
  // Chrome on an iPad (CriOS) is WebKit.
  return /Chrome\/|Chromium\/|Edg\//.test(agent) ? "chromium" : "webkit";
}

const rows = new Map<string, readonly LayoutSnapshotRow[]>();
const loading = new Map<string, Promise<void>>();
const waited = new Map<string, number>();

/** How long an opening holds for the layouts it has asked for. */
const LONGEST_WAIT_MS = 2500;

function snapshotsPath(pageId: string, part?: number): string {
  return `/api/v1/pages/${encodeURIComponent(pageId)}/layout-snapshots?layout=${encodeURIComponent(LAYOUT_CODE_VERSION)}&engine=${layoutEngineName()}${part === undefined ? "" : `&part=${part}`}`;
}

/**
 * Read the layouts kept for this page (again: a reopening meets what another
 * reader kept since). The body's first, on its own: the opening holds for that
 * alone, and the rest (a book's are over a megabyte) comes behind it for the
 * parts that follow.
 */
/** When each page's read was last asked: an opening asks as it starts and
 * again once its content is in, and the second takes the first's. */
const askedAt = new Map<string, number>();
const SAME_OPENING_MS = 20_000;

export function preloadLayoutSnapshots(pageId: string): Promise<void> {
  if (!LAYOUT_CODE_VERSION || !pageId) return Promise.resolve();
  const pending = loading.get(pageId);
  if (pending) return pending;
  if (Date.now() - (askedAt.get(pageId) ?? -Infinity) < SAME_OPENING_MS && rows.has(pageId)) return Promise.resolve();
  askedAt.set(pageId, Date.now());
  waited.delete(pageId);
  rows.set(pageId, []);
  const read = apiGet<{ rows?: LayoutSnapshotRow[]; accepts?: unknown }>(snapshotsPath(pageId, 0)).then((answer) => {
    noteAccepted(answer?.accepts);
    rows.set(pageId, Array.isArray(answer?.rows) ? answer.rows : []);
  }).catch(() => {
    rows.set(pageId, []);
  }).finally(() => {
    loading.delete(pageId);
    void apiGet<{ rows?: LayoutSnapshotRow[]; accepts?: unknown }>(snapshotsPath(pageId)).then((answer) => {
      noteAccepted(answer?.accepts);
      if (Array.isArray(answer?.rows)) rows.set(pageId, answer.rows);
    }).catch(() => undefined);
  });
  loading.set(pageId, read);
  return read;
}

/** Whether an opening should still hold for the read under way: for
 * LONGEST_WAIT_MS from the moment it first holds (the read is asked as the
 * page is, and a slow page read is no reason to give up on the rows). */
export function layoutSnapshotsAwaited(pageId: string): boolean {
  if (!loading.has(pageId)) return false;
  const now = Date.now();
  if (!waited.has(pageId)) waited.set(pageId, now);
  return now - waited.get(pageId)! < LONGEST_WAIT_MS;
}

export function layoutSnapshotRows(pageId: string, part: number): readonly LayoutSnapshotRow[] {
  return (rows.get(pageId) ?? []).filter((row) => row.part === part && row.layout === LAYOUT_CODE_VERSION && row.engine === layoutEngineName());
}

export function forgetLayoutSnapshots(pageId: string): void {
  rows.delete(pageId);
  askedAt.delete(pageId);
}

export interface LayoutSnapshotPartOut {
  readonly part: number;
  readonly segment?: number;
  readonly segments?: number;
  readonly firstBlock?: number;
  readonly captureId?: string;
  readonly contentHash: string;
  readonly blockCount: number;
  readonly continues: boolean;
  readonly pages: readonly (readonly ProjectionNodeRecipe[])[];
  readonly measures: readonly (readonly (readonly number[] | null)[])[];
  readonly lines?: readonly (readonly (readonly number[] | null)[])[];
  readonly autoLines?: readonly (readonly (readonly unknown[] | null)[])[];
}

/* ---- What a part may carry besides its pages ------------------------------ */




export type LayoutRowExtra = "lines" | "autoLines";
const EXTRAS: readonly LayoutRowExtra[] = ["lines", "autoLines"];
let serverAccepts: ReadonlySet<LayoutRowExtra> = new Set();
const printAccepts = new Map<string, ReadonlySet<LayoutRowExtra>>();

function acceptedOf(value: unknown): ReadonlySet<LayoutRowExtra> {
  return new Set(Array.isArray(value) ? EXTRAS.filter((extra) => value.includes(extra)) : []);
}

/** The latest answer says: a server rolled back says none. */
function noteAccepted(value: unknown): void {
  serverAccepts = acceptedOf(value);
}

/** Whether rows written for this page may carry `extra`. */
export function layoutRowExtraAccepted(pageId: string, extra: LayoutRowExtra): boolean {
  return (printAccepts.get(pageId) ?? serverAccepts).has(extra);
}

/** A part without its extras. */
function plainPart(part: LayoutSnapshotPartOut): LayoutSnapshotPartOut {
  const plain: { -readonly [K in keyof LayoutSnapshotPartOut]?: LayoutSnapshotPartOut[K] } = { ...part };
  delete plain.lines;
  delete plain.autoLines;
  return plain as LayoutSnapshotPartOut;
}

/** Tests only. */
export function setLayoutRowExtrasAcceptedForTests(value: unknown): void {
  noteAccepted(value);
  printAccepts.clear();
}

/** Keep these parts of the layout this reader reached. */
export async function storeLayoutSnapshots(pageId: string, fontsKey: string, parts: readonly LayoutSnapshotPartOut[], partCount: number): Promise<boolean> {
  if (!LAYOUT_CODE_VERSION || !parts.length) return false;
  const put = (sent: readonly LayoutSnapshotPartOut[]) => apiPut(`/api/v1/pages/${encodeURIComponent(pageId)}/layout-snapshots`, {
    layout: LAYOUT_CODE_VERSION, engine: layoutEngineName(), fontsKey, parts: sent, partCount,
  });
  try {
    await put(parts);
    return true;
  } catch {
    // A server that said it keeps the extras and then refused (another of its
    // instances, older, mid-release): the pages are kept without them, and
    // nothing is sent with them again until a server says so anew.
    if (!parts.some((part) => part.lines || part.autoLines)) return false;
    serverAccepts = new Set();
    printAccepts.delete(pageId);
    try {
      await put(parts.map(plainPart));
      return true;
    } catch {
      return false;
    }
  }
}

export function traceLayoutSnapshot(entry: string): void {
  if (typeof window === "undefined") return;
  const sink = (window as Window & { __docsLayoutSnapshotTrace?: unknown[] }).__docsLayoutSnapshotTrace;
  if (Array.isArray(sink)) sink.push(`${Math.round(performance.now())} ${entry}`);
}

const documentPartOwners = new Map<string, { owners: () => ReadonlyMap<string, number> }>();

/** Which part holds each loaded block of a page held as parts (the parts hook
 * says, for as long as it holds the page). */
export function holdDocumentPartOwners(pageId: string, owners: () => ReadonlyMap<string, number>): () => void {
  const held = { owners };
  documentPartOwners.set(pageId, held);
  return () => {
    if (documentPartOwners.get(pageId) === held) documentPartOwners.delete(pageId);
  };
}

export function documentPartOwnersOf(pageId: string): ReadonlyMap<string, number> | null {
  return documentPartOwners.get(pageId)?.owners() ?? null;
}

 

const printRows = new Map<string, readonly LayoutSnapshotRow[]>();
const printParts = new Map<string, { readonly blocks: readonly (readonly unknown[])[]; readonly attrs: unknown }>();

export function handPrintLayouts(
  pageId: string,
  layouts: { readonly rows?: readonly LayoutSnapshotRow[]; readonly part_blocks?: readonly number[]; readonly accepts?: unknown } | null | undefined,
  content: unknown,
): void {
  printRows.delete(pageId);
  printParts.delete(pageId);
  printAccepts.delete(pageId);
  if (!layouts) return;
  // The server printing says what the rows it keeps from this print may carry.
  printAccepts.set(pageId, acceptedOf(layouts.accepts));
  if (layouts.rows?.length) printRows.set(pageId, layouts.rows);
  const flat = (content as { content?: unknown; attrs?: unknown } | null)?.content;
  const sizes = layouts.part_blocks;
  if (!Array.isArray(flat) || !sizes?.length) return;
  const blocks: unknown[][] = [];
  let at = 0;
  for (const size of sizes) {
    blocks.push(flat.slice(at, at + size));
    at += size;
  }
  // The server's count of what it merged: anything else names no parts.
  if (at === flat.length) printParts.set(pageId, { blocks, attrs: (content as { attrs?: unknown }).attrs });
  traceLayoutSnapshot(`print handed ${layouts.rows?.length ?? 0} rows, ${sizes.length} parts of ${at} blocks for ${flat.length}`);
}

export function printRowsOf(pageId: string): readonly LayoutSnapshotRow[] | null {
  return printRows.get(pageId) ?? null;
}

export function printPartsOf(pageId: string): { readonly blocks: readonly (readonly unknown[])[]; readonly attrs: unknown } | null {
  return printParts.get(pageId) ?? null;
}
