"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// ---------------------------------------------------------------------------
// Memory-evidence "view source" locate protocol
//
// The memory detail drawer's "View source" (NodeDetailDrawer) opens the
// source page as a note window and asks the page's editor to scroll to +
// flash the quoted passage. Two delivery paths cover the two window states:
//
//   1. Window already open — the mounted NoteEditor hears the
//      LOCATE_EVIDENCE_EVENT CustomEvent (scoped by pageId in the detail)
//      and locates immediately.
//   2. Window just opened — the editor mounts and hydrates AFTER the event
//      fired, so the request is ALSO parked in a pageId-keyed pending buffer
//      the editor consumes once its content hydration completes.
//
// A consumed or expired request is dropped, so a stale locate can never fire
// on a later unrelated mount (TTL below). Locating is best-effort by design:
// when the quote no longer matches (the document was edited) the editor
// silently degrades to just showing the opened page — no error surface.
// ---------------------------------------------------------------------------

export const LOCATE_EVIDENCE_EVENT = "mrnote:locate-evidence";

/** How long a parked locate request stays valid before it is dropped. */
const PENDING_LOCATE_TTL_MS = 30_000;

export interface EvidenceLocateRequest {
  /** Target page whose editor should locate the quote. */
  pageId: string;
  /** The evidence's quoted source excerpt (primary locator). */
  quoteText: string;
  /**
   * Optional anchor hints recorded at evidence-write time. The offsets are
   * PLAIN-TEXT offsets into the source snapshot — the editor treats them as
   * a disambiguation hint (nearest match wins), never as an exact position,
   * since the document may have been edited since.
   */
  documentNodeId?: string | null;
  startOffset?: number | null;
  endOffset?: number | null;
}

/** The provenance fields of a memory evidence row a "view source" reads. */
export interface EvidenceSourceFields {
  source_type: string;
  metadata_json?: Record<string, unknown> | null;
  source_anchor_json?: Record<string, unknown> | null;
}

/**
 * The source page a "view source" jump should open, or null when the evidence
 * does not trace back to a notebook page. Prefers an explicit
 * `source_anchor_json.page_id` (evidence-v2 anchors), then the pipeline's
 * `metadata_json.source_ref` for `notebook_page` evidence — skipping the
 * ad-hoc selection case where source_ref is the notebook itself.
 */
export function resolveEvidencePageId(
  evidence: EvidenceSourceFields,
  notebookId: string,
): string | null {
  const anchor = evidence.source_anchor_json;
  const anchorPageId =
    anchor && typeof anchor.page_id === "string" ? anchor.page_id.trim() : "";
  if (anchorPageId && anchorPageId !== notebookId) return anchorPageId;
  if (evidence.source_type !== "notebook_page") return null;
  const meta = evidence.metadata_json;
  const ref =
    meta && typeof meta.source_ref === "string" ? meta.source_ref.trim() : "";
  if (!ref || ref === notebookId) return null;
  return ref;
}

interface PendingLocateEntry {
  request: EvidenceLocateRequest;
  expiresAt: number;
}

const pendingByPageId = new Map<string, PendingLocateEntry>();

/**
 * Park a locate request for `request.pageId` and broadcast it to any already
 * mounted editor. Call right after opening (or focusing) the page window.
 */
export function requestEvidenceLocate(request: EvidenceLocateRequest): void {
  if (typeof window === "undefined") return;
  const pageId = request.pageId?.trim();
  if (!pageId || !request.quoteText?.trim()) return;
  pendingByPageId.set(pageId, {
    request,
    expiresAt: Date.now() + PENDING_LOCATE_TTL_MS,
  });
  window.dispatchEvent(
    new CustomEvent<EvidenceLocateRequest>(LOCATE_EVIDENCE_EVENT, {
      detail: request,
    }),
  );
}

/**
 * Claim (and clear) the pending locate request for a page, if one is parked
 * and still fresh. Consume-once: a second caller gets `null`, so the live
 * event listener and the post-hydration consumer can never double-fire.
 */
export function consumePendingEvidenceLocate(
  pageId: string,
): EvidenceLocateRequest | null {
  const key = pageId?.trim();
  if (!key) return null;
  const entry = pendingByPageId.get(key);
  if (!entry) return null;
  pendingByPageId.delete(key);
  if (entry.expiresAt < Date.now()) return null;
  return entry.request;
}
