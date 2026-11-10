/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { apiGet } from "@mrnote/core/web/lib/api.ts";

export interface PdfAnnotationReadItem {
  readonly id: string;
  readonly type: string;
  readonly payload_json: Record<string, unknown>;
  readonly created_at?: string | null;
  readonly created_by?: string | null;
  readonly payload_sha256?: string | null;
}

interface PdfAnnotationReadResponse {
  readonly items?: unknown;
}

export type PdfAnnotationReadScope = Readonly<Record<never, never>>;

interface PdfAnnotationRead {
  readonly promise: Promise<readonly PdfAnnotationReadItem[]>;
  /** The data item's write generation the read began under. */
  readonly writeGeneration: number;
}

const inFlightReads = new Map<string, PdfAnnotationRead>();
const scopedReads = new WeakMap<
  PdfAnnotationReadScope,
  Map<string, PdfAnnotationRead>
>();
const writeGenerations = new Map<string, number>();

/**
 * A reader annotation of `dataItemId` was written: a snapshot read before
 * the write is stale, and the next read of the item fetches again. A
 * scope keeps its snapshot for the whole open document, so without this a
 * re-read after a save (whose new revision re-hydrates) brought the old
 * record back over the edit just made (user audit PDF 0016).
 */
export function notePdfAnnotationWrite(dataItemId: string): void {
  const next = (writeGenerations.get(dataItemId) ?? 0) + 1;
  writeGenerations.set(dataItemId, Number.isSafeInteger(next) ? next : 1);
}
interface MountedPdfAnnotationScope {
  readonly scope: PdfAnnotationReadScope;
  retainCount: number;
  disposeTimer: ReturnType<typeof setTimeout> | null;
}

const mountedScopes = new Map<string, MountedPdfAnnotationScope>();

function schedulePdfAnnotationScopeDisposal(
  openIdentity: string,
  entry: MountedPdfAnnotationScope,
): void {
  if (entry.retainCount > 0 || entry.disposeTimer) return;
  const timer = setTimeout(() => {
    const latest = mountedScopes.get(openIdentity);
    if (
      latest === entry &&
      latest.retainCount === 0 &&
      latest.disposeTimer === timer
    ) {
      mountedScopes.delete(openIdentity);
    }
  }, 0);
  entry.disposeTimer = timer;
}

export function createPdfAnnotationReadScope(): PdfAnnotationReadScope {
  return Object.freeze({});
}

export function getPdfAnnotationReadScope(
  openIdentity: string,
): PdfAnnotationReadScope {
  const current = mountedScopes.get(openIdentity);
  if (current) return current.scope;
  const scope = createPdfAnnotationReadScope();
  const entry = {
    scope,
    retainCount: 0,
    disposeTimer: null,
  };
  mountedScopes.set(openIdentity, entry);
  // Concurrent React can abandon a render before its effect retains this
  // entry. Start zero-reference disposal immediately; a committed mount
  // cancels it in retainPdfAnnotationReadScope.
  schedulePdfAnnotationScopeDisposal(openIdentity, entry);
  return scope;
}

/** Keep one open-window scope stable across React StrictMode's immediate
 * cleanup/remount cycle, while releasing it after a real close. */
export function retainPdfAnnotationReadScope(openIdentity: string): () => void {
  const scope = getPdfAnnotationReadScope(openIdentity);
  const entry = mountedScopes.get(openIdentity);
  if (!entry || entry.scope !== scope) return () => undefined;
  if (entry.disposeTimer) clearTimeout(entry.disposeTimer);
  entry.disposeTimer = null;
  entry.retainCount += 1;
  let released = false;
  return () => {
    if (released) return;
    released = true;
    const latest = mountedScopes.get(openIdentity);
    if (!latest || latest.scope !== scope) return;
    latest.retainCount = Math.max(0, latest.retainCount - 1);
    schedulePdfAnnotationScopeDisposal(openIdentity, latest);
  };
}

function isPdfAnnotationReadItem(
  value: unknown,
): value is PdfAnnotationReadItem {
  if (!value || typeof value !== "object") return false;
  const item = value as Partial<PdfAnnotationReadItem>;
  return (
    typeof item.id === "string" &&
    item.id.length > 0 &&
    typeof item.type === "string" &&
    item.type.length > 0 &&
    Boolean(item.payload_json) &&
    typeof item.payload_json === "object" &&
    !Array.isArray(item.payload_json)
  );
}

function readAllPdfAnnotations(
  dataItemId: string,
  options: {
    readonly scope?: PdfAnnotationReadScope;
    readonly refresh?: boolean;
  },
): Promise<readonly PdfAnnotationReadItem[]> {
  const reads = options.scope
    ? (scopedReads.get(options.scope) ?? new Map())
    : inFlightReads;
  if (options.scope && !scopedReads.has(options.scope)) {
    scopedReads.set(options.scope, reads);
  }
  const writeGeneration = writeGenerations.get(dataItemId) ?? 0;
  const current = reads.get(dataItemId);
  if (
    current &&
    !options.refresh &&
    current.writeGeneration === writeGeneration
  ) {
    return current.promise;
  }

  const pending = apiGet<PdfAnnotationReadResponse>(
    `/api/v1/data-items/${dataItemId}/annotations`,
  ).then((response) => {
    if (!Array.isArray(response.items)) {
      throw new Error("Invalid PDF annotation response");
    }
    if (!response.items.every(isPdfAnnotationReadItem)) {
      throw new Error("Invalid PDF annotation item");
    }
    return response.items;
  });
  const read: PdfAnnotationRead = { promise: pending, writeGeneration };
  reads.set(dataItemId, read);
  void pending.then(
    () => {
      // A component-owned scope retains the accepted snapshot for the whole
      // open document. Unscoped callers keep the historical in-flight-only
      // behavior so a later independent read cannot inherit stale state.
      if (!options.scope && reads.get(dataItemId) === read) {
        reads.delete(dataItemId);
      }
    },
    () => {
      if (reads.get(dataItemId) === read) {
        reads.delete(dataItemId);
      }
    },
  );
  return pending;
}

/**
 * Coalesce the PDF reader's same-open annotation hydration burst into one
 * workspace-scoped request. A component-owned scope retains the accepted
 * snapshot until that document closes; explicit authority reconciliation uses
 * `refresh` and replaces it with a new server result.
 */
export async function readPdfAnnotationsByType(
  dataItemId: string,
  type: string,
  options: {
    readonly scope?: PdfAnnotationReadScope;
    readonly refresh?: boolean;
  } = {},
): Promise<{ readonly items: readonly PdfAnnotationReadItem[] }> {
  const items = await readAllPdfAnnotations(dataItemId, options);
  return {
    items: items.filter((item) => item.type === type),
  };
}
