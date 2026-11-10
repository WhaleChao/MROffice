/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * 删除隐藏信息 (ONLYOFFICE's Remove Hidden Information, Acrobat's Remove
 * Hidden Information / Sanitize): what a PDF carries that its pages do not
 * show, found by kind and removed by the kinds chosen.
 *
 * The removal writes the whole file again: an appended revision keeps every
 * earlier one, and with it what was "removed". Objects nothing refers to any
 * more are dropped, so a removed attachment's bytes do not stay behind.
 */
import {
  PDFArray,
  PDFDict,
  PDFDocument,
  PDFName,
  PDFRef,
  PDFStream,
  type PDFObject,
} from "pdf-lib";

import { loadPdfForMutation } from "./encrypted-guard";

export type PdfHiddenInformationKind =
  | "metadata"
  | "attachments"
  | "javascript"
  | "comments"
  | "bookmarks"
  | "thumbnails"
  | "privateData";

export const PDF_HIDDEN_INFORMATION_KINDS: readonly PdfHiddenInformationKind[] = [
  "metadata",
  "attachments",
  "javascript",
  "comments",
  "bookmarks",
  "thumbnails",
  "privateData",
];

/** How many of each kind the file holds (0: none to remove). */
export type PdfHiddenInformationReport = Readonly<Record<PdfHiddenInformationKind, number>>;

const N = (name: string) => PDFName.of(name);

/** Annotations that are not comments: they belong to the page's content or
 * its form, and stay. */
const NOT_COMMENTS = new Set(["/Link", "/Widget", "/Popup", "/PrinterMark", "/TrapNet", "/Watermark"]);

function lookupDict(doc: PDFDocument, value: PDFObject | undefined): PDFDict | undefined {
  const resolved = value instanceof PDFRef ? doc.context.lookup(value) : value;
  return resolved instanceof PDFDict ? resolved : undefined;
}

function lookupArray(doc: PDFDocument, value: PDFObject | undefined): PDFArray | undefined {
  const resolved = value instanceof PDFRef ? doc.context.lookup(value) : value;
  return resolved instanceof PDFArray ? resolved : undefined;
}

function nameOf(doc: PDFDocument, value: PDFObject | undefined): string | undefined {
  const resolved = value instanceof PDFRef ? doc.context.lookup(value) : value;
  return resolved instanceof PDFName ? resolved.asString() : undefined;
}

/** Entries of a name tree (leaves and kids), at most `limit` deep. */
function countNameTree(doc: PDFDocument, root: PDFDict | undefined, depth = 0): number {
  if (!root || depth > 32) return 0;
  let count = 0;
  const names = lookupArray(doc, root.get(N("Names")));
  if (names) count += Math.floor(names.size() / 2);
  const kids = lookupArray(doc, root.get(N("Kids")));
  if (kids) {
    for (let index = 0; index < kids.size(); index += 1) {
      count += countNameTree(doc, lookupDict(doc, kids.get(index)), depth + 1);
    }
  }
  return count;
}

function isJavaScriptAction(doc: PDFDocument, value: PDFObject | undefined): boolean {
  const action = lookupDict(doc, value);
  return Boolean(action && nameOf(doc, action.get(N("S"))) === "/JavaScript");
}

function pageAnnotations(doc: PDFDocument): Array<{ page: PDFDict; annots: PDFArray }> {
  const out: Array<{ page: PDFDict; annots: PDFArray }> = [];
  for (const page of doc.getPages()) {
    const annots = lookupArray(doc, page.node.get(N("Annots")));
    if (annots) out.push({ page: page.node, annots });
  }
  return out;
}

function countOutline(doc: PDFDocument): number {
  const outlines = lookupDict(doc, doc.catalog.get(N("Outlines")));
  if (!outlines) return 0;
  let count = 0;
  const seen = new Set<PDFDict>();
  const walk = (first: PDFObject | undefined, depth: number) => {
    let item = lookupDict(doc, first);
    while (item && !seen.has(item) && depth < 64 && count < 100_000) {
      seen.add(item);
      count += 1;
      walk(item.get(N("First")), depth + 1);
      item = lookupDict(doc, item.get(N("Next")));
    }
  };
  walk(outlines.get(N("First")), 0);
  return count;
}

function inspectDocument(doc: PDFDocument): PdfHiddenInformationReport {
  const catalog = doc.catalog;
  const names = lookupDict(doc, catalog.get(N("Names")));
  let metadata = 0;
  const info = lookupDict(doc, doc.context.trailerInfo.Info);
  if (info) metadata += info.keys().length;
  if (catalog.get(N("Metadata"))) metadata += 1;

  let attachments = countNameTree(doc, lookupDict(doc, names?.get(N("EmbeddedFiles"))));
  let javascript = countNameTree(doc, lookupDict(doc, names?.get(N("JavaScript"))));
  if (isJavaScriptAction(doc, catalog.get(N("OpenAction")))) javascript += 1;
  if (catalog.get(N("AA"))) javascript += 1;
  let comments = 0;
  let thumbnails = 0;
  let privateData = catalog.get(N("PieceInfo")) ? 1 : 0;
  for (const page of doc.getPages()) {
    if (page.node.get(N("AA"))) javascript += 1;
    if (page.node.get(N("Thumb"))) thumbnails += 1;
    if (page.node.get(N("PieceInfo"))) privateData += 1;
    if (page.node.get(N("Metadata"))) metadata += 1;
  }
  for (const { annots } of pageAnnotations(doc)) {
    for (let index = 0; index < annots.size(); index += 1) {
      const annot = lookupDict(doc, annots.get(index));
      if (!annot) continue;
      const subtype = nameOf(doc, annot.get(N("Subtype"))) ?? "";
      if (subtype === "/FileAttachment") attachments += 1;
      else if (!NOT_COMMENTS.has(subtype)) comments += 1;
      if (isJavaScriptAction(doc, annot.get(N("A"))) || annot.get(N("AA"))) javascript += 1;
    }
  }
  return {
    metadata,
    attachments,
    javascript,
    comments,
    bookmarks: countOutline(doc),
    thumbnails,
    privateData,
  };
}

/** What hidden information the file holds, by kind. */
export async function inspectPdfHiddenInformation(bytes: Uint8Array): Promise<PdfHiddenInformationReport> {
  const doc = await loadPdfForMutation(bytes, { updateMetadata: false });
  return inspectDocument(doc);
}

/** Drop every indirect object nothing reachable from the trailer refers to
 * (pdf-lib writes every object it holds, referred to or not). */
function pruneUnreachable(doc: PDFDocument): void {
  const context = doc.context;
  const reachable = new Set<string>();
  const pending: PDFObject[] = [];
  if (context.trailerInfo.Root) pending.push(context.trailerInfo.Root);
  if (context.trailerInfo.Info) pending.push(context.trailerInfo.Info);
  while (pending.length > 0) {
    const value = pending.pop();
    if (value instanceof PDFRef) {
      const key = value.toString();
      if (reachable.has(key)) continue;
      reachable.add(key);
      const target = context.lookup(value);
      if (target) pending.push(target);
    } else if (value instanceof PDFDict) {
      for (const [, entry] of value.entries()) pending.push(entry);
    } else if (value instanceof PDFStream) {
      for (const [, entry] of value.dict.entries()) pending.push(entry);
    } else if (value instanceof PDFArray) {
      for (let index = 0; index < value.size(); index += 1) pending.push(value.get(index));
    }
  }
  for (const [ref] of context.enumerateIndirectObjects()) {
    if (!reachable.has(ref.toString())) context.delete(ref);
  }
}

function removeKinds(doc: PDFDocument, kinds: ReadonlySet<PdfHiddenInformationKind>): void {
  const catalog = doc.catalog;
  const names = lookupDict(doc, catalog.get(N("Names")));
  if (kinds.has("metadata")) {
    // An empty Info: no title, author, dates or producer remain.
    doc.context.trailerInfo.Info = doc.context.register(doc.context.obj({}));
    catalog.delete(N("Metadata"));
    for (const page of doc.getPages()) page.node.delete(N("Metadata"));
  }
  if (kinds.has("attachments")) names?.delete(N("EmbeddedFiles"));
  if (kinds.has("javascript")) {
    names?.delete(N("JavaScript"));
    if (isJavaScriptAction(doc, catalog.get(N("OpenAction")))) catalog.delete(N("OpenAction"));
    catalog.delete(N("AA"));
    for (const page of doc.getPages()) page.node.delete(N("AA"));
  }
  if (kinds.has("bookmarks")) {
    catalog.delete(N("Outlines"));
    if (nameOf(doc, catalog.get(N("PageMode"))) === "/UseOutlines") catalog.delete(N("PageMode"));
  }
  if (kinds.has("thumbnails")) {
    for (const page of doc.getPages()) page.node.delete(N("Thumb"));
  }
  if (kinds.has("privateData")) {
    catalog.delete(N("PieceInfo"));
    for (const page of doc.getPages()) page.node.delete(N("PieceInfo"));
  }
  for (const { page, annots } of pageAnnotations(doc)) {
    const kept: PDFObject[] = [];
    for (let index = 0; index < annots.size(); index += 1) {
      const raw = annots.get(index);
      const annot = lookupDict(doc, raw);
      if (!annot) continue;
      const subtype = nameOf(doc, annot.get(N("Subtype"))) ?? "";
      if (kinds.has("attachments") && subtype === "/FileAttachment") continue;
      if (kinds.has("comments") && !NOT_COMMENTS.has(subtype)) continue;
      // A comment's popup goes with it.
      if (kinds.has("comments") && subtype === "/Popup") {
        const parent = lookupDict(doc, annot.get(N("Parent")));
        const parentType = parent ? nameOf(doc, parent.get(N("Subtype"))) ?? "" : "";
        if (!parent || !NOT_COMMENTS.has(parentType)) continue;
      }
      if (kinds.has("javascript")) {
        if (isJavaScriptAction(doc, annot.get(N("A")))) annot.delete(N("A"));
        annot.delete(N("AA"));
      }
      kept.push(raw);
    }
    if (kept.length === 0) page.delete(N("Annots"));
    else page.set(N("Annots"), doc.context.obj(kept));
  }
  if (names && names.keys().length === 0) catalog.delete(N("Names"));
}

/**
 * The file written again without the kinds chosen: what is left reads as
 * the same pages, and none of what was removed stays in it.
 */
export async function removePdfHiddenInformation(
  bytes: Uint8Array,
  kinds: readonly PdfHiddenInformationKind[],
): Promise<Uint8Array> {
  const doc = await loadPdfForMutation(bytes, { updateMetadata: false });
  removeKinds(doc, new Set(kinds));
  pruneUnreachable(doc);
  return doc.save({ useObjectStreams: false, updateFieldAppearances: false });
}
