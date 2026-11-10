/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { Node as PmNode } from "@tiptap/pm/model";
import { TextSelection, type Selection, type Transaction } from "@tiptap/pm/state";
import { canonicalProjectionDocument, stableProjectionJson } from "@mrnote/core/web/components/console/editor/native-document/paged/paginationProjectionReceipt.ts";
import { isRowSliceTail } from "@mrnote/core/web/components/console/editor/native-document/paged/atomRowSlices.ts";

const sourceFingerprints = new WeakMap<PmNode, string | null>();
const sourceIdentityCounts = new WeakMap<PmNode, Map<string, number>>();

function sourceIdentity(source: unknown): string | null {
  if (!source || typeof source !== "object" || Array.isArray(source)) return null;
  for (const key of ["nodeId", "sourceAnchor"] as const) {
    const value = (source as Record<string, unknown>)[key];
    if (typeof value === "string" && value.trim().length > 0) return `${key}:${value}`;
  }
  return null;
}

function uniqueCanonicalSourceIdentity(doc: PmNode, parent: PmNode): string | null {
  const identity = sourceIdentity(parent.attrs.source);
  if (identity === null) return null;
  let counts = sourceIdentityCounts.get(doc);
  if (!counts) {
    counts = new Map();
    const walk = (node: ReturnType<PmNode["toJSON"]>) => {
      const key = sourceIdentity(node.attrs?.source);
      if (key !== null) counts!.set(key, (counts!.get(key) ?? 0) + 1);
      for (const child of node.content ?? []) walk(child);
    };
    // Real continuations merge to one owner; copied provenance on distinct
    // authored paragraphs remains duplicated and cannot identify a selection.
    walk(canonicalProjectionDocument(doc.toJSON()));
    sourceIdentityCounts.set(doc, counts);
  }
  return counts.get(identity) === 1 ? identity : null;
}

/** Exact source comparison, not a digest granting authority. Schema defaults
 * normalize attrs omitted by a continuation merger without discarding native
 * source identities, authored formatting, marks or non-pagination attrs. */
export function canonicalDocxExportSourceFingerprint(doc: PmNode): string | null {
  if (sourceFingerprints.has(doc)) return sourceFingerprints.get(doc)!;
  try {
    const canonical = canonicalProjectionDocument(doc.toJSON());
    const fingerprint = stableProjectionJson({
      ...canonical,
      content: canonical.content?.map(node => doc.type.schema.nodeFromJSON(node).toJSON()),
    });
    sourceFingerprints.set(doc, fingerprint);
    return fingerprint;
  } catch { sourceFingerprints.set(doc, null); return null; }
}

/** Text positions are counted in source leaves rather than page wrappers.
 * Repeated table headers and empty continuation bridges carry no new source.
 * Side-store note slices have no text-position proof here and fail closed. */
function sourceTextEndpoint(doc: PmNode, position: number): string | null {
  let units = 0, endpoint: number | null = null, unknown = false;
  const resolved = doc.resolve(position), parent = resolved.parent;
  const parentStart = resolved.start();
  let spanStart: number | null = null, spanEnd: number | null = null;
  const walk = (node: PmNode, start: number): void => {
    if (spanEnd !== null || unknown) return;
    const attrs = node.attrs;
    if (attrs.pageTableRepeatedHeader === true || attrs.pageTableVerticalContinuation === true ||
        attrs.pageTableCellEmptyHead === true || attrs.pageTableCellEmptyTail === true ||
        attrs.pageSplitListItemBridge === true) return;
    if (node.type.name === "footnoteList" && (attrs.noteSliceStart != null || attrs.noteSliceEnd != null)) { unknown = true; return; }
    // A generated list the page cut between two rows is one source atom:
    // the copies after the first carry no source of their own.
    if (isRowSliceTail(node)) return;
    if (node.isText) {
      if (position >= start && position <= start + node.nodeSize) endpoint = units + position - start;
      units += node.nodeSize;
      return;
    }
    if (node.isAtom) { units += 1; return; }
    const selectedParent = start + 1 === parentStart && node === parent;
    if (selectedParent) spanStart = units;
    node.forEach((child, offset) => walk(child, start + 1 + offset));
    if (selectedParent) spanEnd = units;
  };
  doc.forEach((child, offset) => walk(child, offset));
  if (unknown || endpoint === null || spanStart === null || spanEnd === null) return null;
  const identity = uniqueCanonicalSourceIdentity(doc, parent);
  return stableProjectionJson({ units: endpoint, type: parent.type.name,
    
    // source-leaf span. Unproved paragraph splitting then fails closed.
    source: identity, span: identity === null ? [spanStart, spanEnd] : null });
}

/** Only a recognized derived pagination transaction plus exact source and
 * source-selection equality can advance a DOCX export's live authority. */
export function preservesCanonicalDocxExportSelection(
  transaction: Transaction,
  previousSelection: Selection,
): boolean {
  // Derived pagination is the reflow transaction itself and the identity
  // back-fill appended to it: a cut block's fragments each carry the head's
  // attrs, and the block-identity plugin re-mints the copies. Both leave the
  // canonical (merged) document untouched, which the checks below prove.
  if ((transaction.getMeta("structuralRebalance") !== true &&
       transaction.getMeta("mrnoteBlockIdentity") !== true) ||
      !transaction.docChanged || !(previousSelection instanceof TextSelection) ||
      !(transaction.selection instanceof TextSelection)) return false;
  if (transaction.steps.some((step, index) => {
    const encoded = step.toJSON();
    return (encoded.stepType === "replace" || encoded.stepType === "replaceAround") &&
      encoded.from === 0 && encoded.to === transaction.docs[index].content.size;
  })) return false;
  const before = canonicalDocxExportSourceFingerprint(transaction.before);
  if (before === null || before !== canonicalDocxExportSourceFingerprint(previousSelection.$anchor.doc) ||
      before !== canonicalDocxExportSourceFingerprint(transaction.doc)) return false;
  const beforeAnchor = sourceTextEndpoint(previousSelection.$anchor.doc, previousSelection.anchor);
  const beforeHead = sourceTextEndpoint(previousSelection.$anchor.doc, previousSelection.head);
  return beforeAnchor !== null && beforeHead !== null &&
    beforeAnchor === sourceTextEndpoint(transaction.doc, transaction.selection.anchor) &&
    beforeHead === sourceTextEndpoint(transaction.doc, transaction.selection.head);
}
