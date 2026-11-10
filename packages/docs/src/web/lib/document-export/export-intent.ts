/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { PAGE_REGION_RECEIPT } from "@mrnote/core/web/components/console/editor/native-document/paged/paginationProjectionKeys.ts";
import { isApiRequestError } from "@mrnote/core/web/lib/api.ts";
import type { PageExportFormat } from "@mrnote/core/web/lib/page-export.ts";

export type DocumentExportFormat = Extract<
  PageExportFormat,
  "pdf" | "docx" | "odt" | "doc" | "rtf" | "dotx" | "wordxml" | "markdown" | "html" | "text"
>;

export type DocumentExportRenderer =
  | "canonical-server"
  | "inline-server"
  | "offline-browser"
  | "browser-print";

export type DocumentExportSource = "saved-page" | "current-editor";
export type DocumentExportRevisionPolicy = "accept" | "preserve";
export type DocumentExportFidelity =
  | "layout"
  | "document"
  | "source-package"
  | "structured-text"
  | "web-structure"
  | "plain-text"
  | "compatibility"
  | "screen-print"
  | "fillable-form";

export interface ExportIntent {
  readonly id: string;
  readonly format: DocumentExportFormat;
  readonly renderer: DocumentExportRenderer;
  readonly source: DocumentExportSource;
  readonly revisionPolicy: DocumentExportRevisionPolicy;
  readonly fidelity: DocumentExportFidelity;
}

function intent(
  renderer: DocumentExportRenderer,
  format: DocumentExportFormat,
  revisionPolicy: DocumentExportRevisionPolicy,
  fidelity: DocumentExportFidelity,
): ExportIntent {
  return Object.freeze({
    id: `${renderer}:${format}:${revisionPolicy}`,
    format,
    renderer,
    source:
      renderer === "canonical-server" ? "saved-page" : "current-editor",
    revisionPolicy,
    fidelity,
  });
}

/**
 * The persisted-page export surface has one authority: the authenticated,
 * CSRF-bound server endpoint. Keeping the menu model here prevents a browser
 * renderer from being reintroduced beside a same-named canonical format.
 */
export const CANONICAL_DOCUMENT_EXPORT_INTENTS = Object.freeze([
  intent("canonical-server", "pdf", "accept", "layout"),
  // Word's Save as PDF > Options > Publish what: "Document showing markup"
  // (the review display's markup printed on the pages) beside "Document".
  intent("canonical-server", "pdf", "preserve", "layout"),
  intent("canonical-server", "docx", "accept", "document"),
  intent("canonical-server", "docx", "preserve", "source-package"),
  // OpenDocument Text, written by the server from the Word export (Word's
  // own Save As .odt; formatting ODF cannot hold may look different).
  intent("canonical-server", "odt", "accept", "compatibility"),
  // Word's Save As Word 97-2003 Document and Rich Text Format, written by the
  // server from the Word export (features the older formats cannot hold may
  // look different), and Save As Word Template (the same package, a template).
  intent("canonical-server", "doc", "accept", "compatibility"),
  intent("canonical-server", "rtf", "accept", "compatibility"),
  intent("canonical-server", "dotx", "accept", "document"),
  // Word's Save As Word XML Document: the same package as one XML file.
  intent("canonical-server", "wordxml", "accept", "document"),
  intent("canonical-server", "markdown", "accept", "structured-text"),
  intent("canonical-server", "html", "accept", "web-structure"),
  intent("canonical-server", "text", "accept", "plain-text"),
]);

/** Save as a fillable PDF form (the Forms tab): the final copy's PDF with an
 *  AcroForm field over each content control (lib/document-export/fillable-pdf). */
export const FILLABLE_PDF_EXPORT_INTENT: ExportIntent = Object.freeze({
  ...intent("canonical-server", "pdf", "accept", "fillable-form"),
  id: "canonical-server:pdf:fillable",
});

export const BROWSER_PRINT_EXPORT_INTENT = intent(
  "browser-print",
  "pdf",
  "accept",
  "screen-print",
);

/** Print with Word's Print Markup checked: the tracked changes printed as the
 * review display shows them (paged/printRevisionMarkup). */
export const BROWSER_PRINT_MARKUP_EXPORT_INTENT = intent(
  "browser-print",
  "pdf",
  "preserve",
  "screen-print",
);

/** Print as Word prints by default: with its markup while the document holds
 * tracked changes (Print Markup is checked whenever it does), else the
 * document. */
export function defaultPrintIntent(hasTrackedChanges: boolean): ExportIntent {
  return hasTrackedChanges ? BROWSER_PRINT_MARKUP_EXPORT_INTENT : BROWSER_PRINT_EXPORT_INTENT;
}

/**
 * Guest/unsaved documents cannot use the canonical endpoint. These operate on
 * the current editor snapshot and never masquerade as the saved-page export.
 * Text formats are written in the browser; a Word file is written by the
 * server's Word writer from the snapshot (`inline-server`, which needs a
 * connection), never in the browser.
 */
export const OFFLINE_DOCUMENT_EXPORT_INTENTS = Object.freeze([
  intent("offline-browser", "markdown", "accept", "structured-text"),
  intent("offline-browser", "html", "accept", "web-structure"),
  intent("offline-browser", "text", "accept", "plain-text"),
  intent("inline-server", "docx", "accept", "document"),
  intent("inline-server", "docx", "preserve", "document"),
  BROWSER_PRINT_MARKUP_EXPORT_INTENT,
  BROWSER_PRINT_EXPORT_INTENT,
]);

/**
 * The rows of a document's Export menu. A saved copy of a Word document keeps its
 * tracked changes until someone accepts or rejects them (Microsoft Support, "Track
 * changes in Word"), and Word's PDF of a document holding some shows its markup
 * ("Document showing markup") unless "Document" is chosen, so a document holding
 * tracked changes offers, per format, the copy that keeps them first and the
 * final copy second, each named as what it is. One holding none has one file of
 * each format to offer: both would say the same.
 */
export function documentExportMenuIntents(
  intents: readonly ExportIntent[],
  hasTrackedChanges: boolean,
): readonly ExportIntent[] {
  const paired = (entry: ExportIntent) => entry.format === "docx" || entry.format === "pdf";
  if (!hasTrackedChanges) {
    return intents.filter((entry) => !(paired(entry) && entry.revisionPolicy === "preserve"));
  }
  const ordered = new Map<string, ExportIntent[]>();
  for (const entry of intents) {
    if (!paired(entry)) continue;
    const group = ordered.get(entry.format) ?? [];
    group.push(entry);
    ordered.set(entry.format, group);
  }
  for (const [format, group] of ordered) {
    ordered.set(format, [
      ...group.filter((entry) => entry.revisionPolicy === "preserve"),
      ...group.filter((entry) => entry.revisionPolicy !== "preserve"),
    ]);
  }
  const next = new Map<string, number>();
  return intents.map((entry) => {
    if (!paired(entry)) return entry;
    const at = next.get(entry.format) ?? 0;
    next.set(entry.format, at + 1);
    return ordered.get(entry.format)![at];
  });
}

export type ExportPhase = "idle" | "preparing" | "ready" | "failed";

export interface ExportStatus {
  readonly phase: ExportPhase;
  readonly intent: ExportIntent | null;
  readonly error: string | null;
}

export const INITIAL_EXPORT_STATUS: ExportStatus = Object.freeze({
  phase: "idle",
  intent: null,
  error: null,
});

export type ExportStatusEvent =
  | { readonly type: "start"; readonly intent: ExportIntent }
  | { readonly type: "succeed"; readonly intentId: string }
  | { readonly type: "fail"; readonly intentId: string; readonly error: string };

/** Ignore stale completions so an older operation cannot overwrite a newer UI. */
export function exportStatusReducer(
  status: ExportStatus,
  event: ExportStatusEvent,
): ExportStatus {
  if (event.type === "start") {
    return { phase: "preparing", intent: event.intent, error: null };
  }
  if (!status.intent || status.intent.id !== event.intentId) return status;
  if (event.type === "succeed") {
    return { phase: "ready", intent: status.intent, error: null };
  }
  return { phase: "failed", intent: status.intent, error: event.error };
}

export function exportErrorMessage(
  error: unknown,
  fallback = "Export failed",
  trustedMessages: readonly string[] = [],
): string {
  const message = error instanceof Error ? error.message.trim() : "";
  return message && trustedMessages.includes(message) ? message : fallback;
}

/** The canonical PDF's typed refusals of content its renderer cannot draw. */
const PDF_EXPORT_CONTENT_REFUSALS: ReadonlySet<string> = new Set([
  "document_pdf_export_unsupported_node",
  "document_pdf_export_unsupported_mark",
  "document_pdf_export_invalid_table",
  "document_pdf_export_invalid_image",
  "document_pdf_export_invalid_image_row",
  "document_pdf_export_invalid_chart",
  "document_pdf_export_layout_overflow",
]);

/** The canonical PDF refused an object its renderer cannot draw (a typed
 * `document_pdf_export_*` refusal of a node, mark, table, picture or chart).
 * Nothing went wrong and nothing changed, so the user is told what to do
 * instead of "the exporter stopped unexpectedly" (live audit WUA 012: a
 * comment's mark read as a crash). */
export function isPdfExportUnsupportedObjectError(error: unknown): boolean {
  return (
    isApiRequestError(error) &&
    PDF_EXPORT_CONTENT_REFUSALS.has(error.code ?? "")
  );
}

/** Saving may retire a cached page projection without changing export content.
 * Keep all document structure, formatting and native identities in the guard;
 * only this root-level, derived pagination receipt is not source authority. */
export function documentExportSourceFingerprint(document: unknown): string {
  if (!document || typeof document !== "object" || Array.isArray(document))
    return JSON.stringify(document) ?? "null";
  const root = document as Record<string, unknown>;
  if (!root.attrs || typeof root.attrs !== "object" || Array.isArray(root.attrs))
    return JSON.stringify(root);
  const attrs = { ...root.attrs } as Record<string, unknown>;
  delete attrs[PAGE_REGION_RECEIPT];
  return JSON.stringify({ ...root, attrs });
}
