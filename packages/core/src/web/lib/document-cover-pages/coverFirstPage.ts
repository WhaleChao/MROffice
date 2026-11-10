/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Word keeps page numbers, headers and footers off a cover page by putting
 * the cover on page 1 of a section whose "Different First Page" (w:titlePg)
 * is on and whose first-page header and footer are empty (measured in Word
 * for Mac: without w:titlePg the cover prints the header and "1"; with it,
 * page 1 prints neither and page 2 prints "2" - the cover still counts).
 *
 * MRAI's automatic page number (`showPageNumbers`) is a document-wide
 * footer; this is the one rule that keeps it off the cover, shared by the
 * paged view and the DOCX exporter so they cannot disagree.
 */

import type { Node as ProseMirrorNode } from "@tiptap/pm/model";

import { normalizeSectionBands } from "../../components/console/editor/native-document/headerFooter";
import { normalizeSectionStoryGraph } from "../../components/console/editor/native-document/sectionStoryGraph";

import { DOCUMENT_COVER_PAGE_NODE } from "./coverPages";

interface JsonLikeNode {
  type?: string;
  attrs?: Record<string, unknown> | null;
  content?: readonly JsonLikeNode[] | null;
}

/** Whether the first section uses a different first-page header/footer. */
export function firstSectionHasTitlePage(attrs: Record<string, unknown>): boolean {
  const graph =
    attrs.sectionStoryGraph !== null && attrs.sectionStoryGraph !== undefined
      ? normalizeSectionStoryGraph(attrs.sectionStoryGraph)
      : null;
  if (graph) return graph.sections[0]?.titlePage === true;
  const bands = normalizeSectionBands(attrs.sectionBands);
  if (bands) return bands[0]?.differentFirstPage === true;
  return attrs.differentFirstPage === true;
}

function leadingBlock(node: JsonLikeNode | null | undefined): JsonLikeNode | null {
  let current = node?.content?.[0] ?? null;
  // The editor wraps pages in pageBlock; storage and export are flat.
  while (current && current.type === "pageBlock") current = current.content?.[0] ?? null;
  return current;
}

/** A JSON document (export side) opens with a cover on a title page. */
export function coverHidesFirstPageNumberJson(doc: JsonLikeNode): boolean {
  return (
    leadingBlock(doc)?.type === DOCUMENT_COVER_PAGE_NODE &&
    firstSectionHasTitlePage(doc.attrs ?? {})
  );
}

/** A live document (paged view) opens with a cover on a title page. */
export function coverHidesFirstPageNumber(doc: ProseMirrorNode): boolean {
  let first: ProseMirrorNode | null = doc.firstChild;
  while (first && first.type.name === "pageBlock") first = first.firstChild;
  return (
    first?.type.name === DOCUMENT_COVER_PAGE_NODE &&
    firstSectionHasTitlePage(doc.attrs as Record<string, unknown>)
  );
}
