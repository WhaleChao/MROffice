/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { apiGet, apiPost } from "./api";

// ─── Source anchor (master spec §2.1) ─────────────────────────────────

export type LocatorKind =
  | "pdf_region"
  | "docx_paragraph"
  | "docx_table"
  | "sheet_range"
  | "media_segment"
  | "image_bbox"
  | "web_dom"
  | "archive_path";

export type Locator =
  | {
      kind: "pdf_region";
      page_number: number;
      bbox: [number, number, number, number] | null;
      text_range: { start: number; end: number } | null;
    }
  | {
      kind: "docx_paragraph";
      paragraph_id: string;
      run_offsets: { start: number; end: number } | null;
    }
  | {
      kind: "docx_table";
      table_id: string;
      cell_range: string;
    }
  | {
      kind: "sheet_range";
      sheet_name: string;
      cell_range: string;
    }
  | {
      kind: "media_segment";
      timestamp_start: number;
      timestamp_end: number | null;
      channel: number | null;
    }
  | {
      kind: "image_bbox";
      bbox: [number, number, number, number];
    }
  | {
      kind: "web_dom";
      url: string;
      css_selector: string | null;
      text_range: { start: number; end: number } | null;
    }
  | {
      kind: "archive_path";
      internal_path: string;
      byte_range: { start: number; end: number } | null;
    };

export interface SourceAnchor {
  resource_id: string;
  resource_version_id: string;
  extractor: string;
  confidence: number;
  kind: LocatorKind;
  locator: Locator;
}

// ─── Node types (master spec §2.5) ────────────────────────────────────

export type NodeType =
  | "doc"
  | "page"
  | "section"
  | "paragraph"
  | "heading"
  | "list"
  | "list_item"
  | "code_block"
  | "blockquote"
  | "horizontal_rule"
  | "frontmatter"
  | "image"
  | "figure"
  | "pdf_page"
  | "pdf_text_block"
  | "pdf_region"
  | "docx_paragraph"
  | "docx_table"
  | "docx_section_break"
  | "annotation";

export type DocumentType =
  | "text_document"
  | "pdf_projection"
  | "docx_projection"
  | "custom_workspace";

export interface DocumentNodeRead {
  id: string;
  document_id: string;
  parent_id: string | null;
  node_type: NodeType;
  position: number;
  content_json: Record<string, unknown>;
  layout_json: Record<string, unknown> | null;
  source_anchor_json: SourceAnchor | Record<string, unknown> | null;
  node_revision: number;
  metadata_json: Record<string, unknown>;
  created_at: string;
  updated_at: string;
}

export interface UniversalDocumentRead {
  id: string;
  workspace_id: string;
  notebook_id: string | null;
  source_resource_id: string | null;
  linked_page_id: string | null;
  document_type: DocumentType;
  title: string;
  root_node_id: string | null;
  schema_version: number;
  revision_no: number;
  metadata_json: Record<string, unknown>;
  created_at: string;
  updated_at: string;
}

// ─── API wrappers ─────────────────────────────────────────────────────

/** Fetch a single universal document by id. Throws ApiRequestError on failure. */
export async function getDocument(
  documentId: string,
): Promise<UniversalDocumentRead> {
  return apiGet<UniversalDocumentRead>(
    `/api/v1/universal-documents/${documentId}`,
  );
}

export interface ListDocumentNodesOptions {
  limit?: number;
  cursor?: number | null;
}

export interface DocumentNodesPage {
  items: DocumentNodeRead[];
  next_cursor: number | null;
  limit?: number | null;
}

/** Fetch one flat node-list page for a document. */
export async function listDocumentNodesPage(
  documentId: string,
  opts: ListDocumentNodesOptions = {},
): Promise<DocumentNodesPage> {
  const params = new URLSearchParams();
  if (opts.limit !== undefined) params.set("limit", String(opts.limit));
  if (opts.cursor !== undefined && opts.cursor !== null) {
    params.set("cursor", String(opts.cursor));
  }
  const qs = params.toString();
  const data = await apiGet<DocumentNodesPage>(
    `/api/v1/universal-documents/${documentId}/nodes${qs ? `?${qs}` : ""}`,
  );
  return {
    items: data.items ?? [],
    next_cursor: data.next_cursor ?? null,
    limit: data.limit ?? opts.limit ?? null,
  };
}

/** Fetch the first flat node-list page for a document. */
export async function listDocumentNodes(
  documentId: string,
  opts: ListDocumentNodesOptions = {},
): Promise<DocumentNodeRead[]> {
  const page = await listDocumentNodesPage(documentId, opts);
  return page.items;
}

export interface CreateFromResourceArgs {
  resource_id: string;
  title?: string;
  notebook_id?: string;
}

/** Create a UniversalDocument projection from an existing Resource (PDF, docx, …). */
export async function createDocumentFromResource(
  args: CreateFromResourceArgs,
): Promise<UniversalDocumentRead> {
  return apiPost<UniversalDocumentRead>(
    `/api/v1/universal-documents/from-resource`,
    args,
  );
}

// ─── Projection → editable notebook page ──────────────────────────────

/**
 * Response from ``POST /universal-documents/{id}/convert-to-page``.
 * The returned page id opens through the normal NotebookPage/NoteEditor
 * window, while document_id lets callers keep the derived UniversalDocument
 * relation if they need it later.
 */
export interface ConvertProjectionToPageResult {
  page_id: string;
  document_id: string;
  page_type: string;
}

export interface ConvertProjectionToPageOptions {
  notebook_id?: string | null;
  title?: string;
}

/**
 * Convert a projection UniversalDocument (PDF/DOCX today) into a real
 * editable notebook page. The projection document remains intact; callers
 * should open the returned page_id with WindowType ``note``.
 */
export async function convertProjectionToPage(
  documentId: string,
  options: ConvertProjectionToPageOptions = {},
): Promise<ConvertProjectionToPageResult> {
  const body: { notebook_id?: string; title?: string } = {};
  if (options.notebook_id) body.notebook_id = options.notebook_id;
  if (options.title !== undefined) body.title = options.title;
  return apiPost<ConvertProjectionToPageResult>(
    `/api/v1/universal-documents/${documentId}/convert-to-page`,
    body,
  );
}
