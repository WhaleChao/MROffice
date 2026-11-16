/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { nativeNoteIdToken } from "./native-note-roundtrip";
import { pdfLibParseSpeed } from "./pdf-lib-pacing";

export type PdfManagedNoteSourceIdentity =
  | { readonly kind: "managed"; readonly id: string; readonly text: string | null }
  | { readonly kind: "blocked"; readonly id?: string };
export type PdfManagedNoteIdentityIndex = ReadonlyMap<string, PdfManagedNoteSourceIdentity>;

export function pdfNoteSourceKey(pageNumber: number, objectRef: string): string {
  return `${pageNumber}:${objectRef}`;
}

function decodeManagedId(name: string): string | null {
  if (!name.startsWith("MRNote:")) return null;
  const token = name.slice(7);
  try {
    const id = decodeURIComponent(token);
    if (!id || /[\u0000-\u001f\u007f]/.test(id) || id.startsWith("pdf-native-") ||
      nativeNoteIdToken(id) !== token) return null;
    return id;
  } catch { return null; }
}

/** Identity hints from the actual displayed PDF, never write permission. Scan
 * all annotation names for collisions, but expose mappings only for /Text. */
export async function readManagedPdfTextNoteIdentities(
  sourceBytes: Uint8Array,
): Promise<PdfManagedNoteIdentityIndex | null> {
  if (!sourceBytes.byteLength || sourceBytes.byteLength > 134_217_728) return null;
  try {
    const { PDFDocument, PDFName, PDFDict, PDFArray, PDFRef, PDFBool, PDFString, PDFHexString } = await import("pdf-lib");
    const document = await PDFDocument.load(sourceBytes, { parseSpeed: pdfLibParseSpeed(), updateMetadata: false });
    if (document.isEncrypted || document.getPageCount() > 10_000) return null;
    const index = new Map<string, PdfManagedNoteSourceIdentity>();
    const claims = new Map<string, string[]>();
    const refs = new Map<string, string[]>();
    let count = 0;
    for (const [pageIndex, page] of document.getPages().entries()) {
      const annotations = page.node.lookupMaybe(PDFName.of("Annots"), PDFArray);
      if (!annotations) continue;
      count += annotations.size();
      if (count > 20_000) return null;
      for (let position = 0; position < annotations.size(); position++) {
        const reference = annotations.get(position);
        const annotation = document.context.lookup(reference);
        if (!(reference instanceof PDFRef) || !(annotation instanceof PDFDict)) continue;
        const ref = `${reference.objectNumber}R${reference.generationNumber || ""}`;
        const key = pdfNoteSourceKey(pageIndex + 1, ref);
        refs.set(ref, [...(refs.get(ref) ?? []), key]);
        const value = annotation.get(PDFName.of("NM"));
        const name = document.context.lookup(value);
        if (!(name instanceof PDFString || name instanceof PDFHexString) || !name.decodeText().startsWith("MRNote:")) continue;
        const opaqueName = name.decodeText();
        // Include other subtypes in collision detection without changing their
        // own importers, source coordinates or writer/history contracts.
        claims.set(opaqueName, [...(claims.get(opaqueName) ?? []), key]);
        if (annotation.get(PDFName.of("Subtype")) !== PDFName.of("Text")) continue;
        const id = decodeManagedId(opaqueName);
        const contents = document.context.lookup(annotation.get(PDFName.of("Contents")));
        const decodedText = contents instanceof PDFString || contents instanceof PDFHexString ? contents.decodeText() : null;
        const text = decodedText !== null && decodedText.length <= 32_000 ? decodedText.replace(/\r\n?/g, "\n").trim() : null;
        const parent = annotation.get(PDFName.of("P"));
        const validPage = parent === undefined || (parent instanceof PDFRef && parent.toString() === page.ref.toString());
        index.set(key, id && validPage && !annotation.has(PDFName.of("IRT")) &&
          annotation.get(PDFName.of("MRNoteManaged")) === PDFBool.True
          ? { kind: "managed", id, text } : { kind: "blocked", ...(id ? { id } : {}) });
      }
    }
    for (const keys of [...claims.values(), ...refs.values()]) {
      if (keys.length > 1) for (const key of keys) {
        const prior = index.get(key);
        if (prior) index.set(key, { kind: "blocked", ...(prior.id ? { id: prior.id } : {}) });
      }
    }
    return index;
  } catch { return null; }
}

interface PdfDataSource { getData(): Promise<Uint8Array>; }
const identitiesByDocument = new WeakMap<PdfDataSource, Promise<PdfManagedNoteIdentityIndex | null>>();

/** One bounded read per PDF.js document incarnation. The caller must still
 * check its active document/render identity after awaiting this result. */
export function managedPdfTextNoteIdentitiesForDocument(
  document: PdfDataSource,
): Promise<PdfManagedNoteIdentityIndex | null> {
  let pending = identitiesByDocument.get(document);
  if (!pending) {
    pending = document.getData().then(readManagedPdfTextNoteIdentities).catch(() => null);
    identitiesByDocument.set(document, pending);
  }
  return pending;
}

/** Only the existing reserved wire format marks an authored managed note.
 * Plain text/page resemblance to a native annotation is not such a marker. */
export function managedPdfNoteIdsFromContent(content: Record<string, unknown> | null | undefined): Set<string> {
  const ids = new Set<string>();
  if (!Array.isArray(content?.annotations)) return ids;
  for (const value of content.annotations) {
    if (!value || typeof value !== "object" || typeof value.id !== "string" || typeof value.text !== "string") continue;
    const token = nativeNoteIdToken(value.id);
    if (token && value.text.startsWith(`[[MRNOTE_NATIVE_NOTE_V1|upsert|${token}]]`)) ids.add(value.id);
  }
  return ids;
}










export async function readManagedPdfInkIdentities(
  sourceBytes: Uint8Array,
): Promise<ReadonlyMap<string, string> | null> {
  return readManagedPdfIdentities(sourceBytes, new Set(["Ink"]), "ink.");
}







export async function readManagedPdfShapeIdentities(
  sourceBytes: Uint8Array,
): Promise<ReadonlyMap<string, string> | null> {
  return readManagedPdfIdentities(
    sourceBytes,
    new Set(["Square", "Circle", "Line", "Polygon", "PolyLine"]),
    "shape.",
  );
}











export async function readManagedPdfTextMarkupIdentities(
  sourceBytes: Uint8Array,
): Promise<ReadonlyMap<string, string> | null> {
  return readManagedPdfIdentities(
    sourceBytes,
    new Set(["Highlight", "Underline", "StrikeOut", "Squiggly"]),
    "",
  );
}

async function readManagedPdfIdentities(
  sourceBytes: Uint8Array,
  subtypes: ReadonlySet<string>,
  kindPrefix: string,
): Promise<ReadonlyMap<string, string> | null> {
  const namePrefix = `MRNote:${kindPrefix}`;
  if (!sourceBytes.byteLength || sourceBytes.byteLength > 134_217_728) return null;
  try {
    const { PDFDocument, PDFName, PDFDict, PDFArray, PDFRef, PDFBool, PDFString, PDFHexString } = await import("pdf-lib");
    const document = await PDFDocument.load(sourceBytes, { parseSpeed: pdfLibParseSpeed(), updateMetadata: false });
    if (document.isEncrypted || document.getPageCount() > 10_000) return null;
    const byRef = new Map<string, string>();
    const claims = new Map<string, number>();
    let count = 0;
    for (const page of document.getPages()) {
      const annotations = page.node.lookupMaybe(PDFName.of("Annots"), PDFArray);
      if (!annotations) continue;
      count += annotations.size();
      if (count > 20_000) return null;
      for (let position = 0; position < annotations.size(); position++) {
        const reference = annotations.get(position);
        const annotation = document.context.lookup(reference);
        if (!(reference instanceof PDFRef) || !(annotation instanceof PDFDict)) continue;
        const name = document.context.lookup(annotation.get(PDFName.of("NM")));
        if (!(name instanceof PDFString || name instanceof PDFHexString)) continue;
        const text = name.decodeText();
        if (!text.startsWith("MRNote:")) continue;
        claims.set(text, (claims.get(text) ?? 0) + 1);
        const subtype = annotation.get(PDFName.of("Subtype"));
        if (
          !(subtype instanceof PDFName) ||
          !subtypes.has(subtype.decodeText()) ||
          annotation.get(PDFName.of("MRNoteManaged")) !== PDFBool.True ||
          !text.startsWith(namePrefix)
        ) {
          continue;
        }
        const token = text.slice(namePrefix.length);
        let id: string;
        try {
          id = decodeURIComponent(token);
        } catch {
          continue;
        }
        if (!id.trim() || nativeNoteIdToken(id) !== token) continue;
        byRef.set(`${reference.objectNumber}R${reference.generationNumber || ""}`, text);
      }
    }
    const identities = new Map<string, string>();
    for (const [ref, text] of byRef) {
      if ((claims.get(text) ?? 0) !== 1) continue;
      identities.set(ref, decodeURIComponent(text.slice(namePrefix.length)));
    }
    return identities;
  } catch { return null; }
}

const inkIdentitiesByDocument = new WeakMap<PdfDataSource, Promise<ReadonlyMap<string, string> | null>>();
const shapeIdentitiesByDocument = new WeakMap<PdfDataSource, Promise<ReadonlyMap<string, string> | null>>();
const textMarkupIdentitiesByDocument = new WeakMap<PdfDataSource, Promise<ReadonlyMap<string, string> | null>>();

/** One bounded read per pdf.js document incarnation, like the ink index. */
export function managedPdfTextMarkupIdentitiesForDocument(
  document: PdfDataSource,
): Promise<ReadonlyMap<string, string> | null> {
  let pending = textMarkupIdentitiesByDocument.get(document);
  if (!pending) {
    pending = document.getData().then(readManagedPdfTextMarkupIdentities).catch(() => null);
    textMarkupIdentitiesByDocument.set(document, pending);
  }
  return pending;
}

/** One bounded read per pdf.js document incarnation, like the ink index. */
export function managedPdfShapeIdentitiesForDocument(
  document: PdfDataSource,
): Promise<ReadonlyMap<string, string> | null> {
  let pending = shapeIdentitiesByDocument.get(document);
  if (!pending) {
    pending = document.getData().then(readManagedPdfShapeIdentities).catch(() => null);
    shapeIdentitiesByDocument.set(document, pending);
  }
  return pending;
}

/** One bounded read per pdf.js document incarnation, like the note index. */
export function managedPdfInkIdentitiesForDocument(
  document: PdfDataSource,
): Promise<ReadonlyMap<string, string> | null> {
  let pending = inkIdentitiesByDocument.get(document);
  if (!pending) {
    pending = document.getData().then(readManagedPdfInkIdentities).catch(() => null);
    inkIdentitiesByDocument.set(document, pending);
  }
  return pending;
}

/**
 * Every MRAI annotation's `/NM` (`MRNote:...`), by pdf.js object ref: which
 * object a held edit's name is on the page (lib/pdf/held-annotation-targets).
 * A name claimed twice names neither object.
 */
export async function readManagedPdfAnnotationNames(
  sourceBytes: Uint8Array,
): Promise<ReadonlyMap<string, string> | null> {
  if (!sourceBytes.byteLength || sourceBytes.byteLength > 134_217_728) return null;
  try {
    const { PDFDocument, PDFName, PDFDict, PDFArray, PDFRef, PDFString, PDFHexString } = await import("pdf-lib");
    const document = await PDFDocument.load(sourceBytes, { parseSpeed: pdfLibParseSpeed(), updateMetadata: false });
    if (document.isEncrypted || document.getPageCount() > 10_000) return null;
    const byRef = new Map<string, string>();
    const claims = new Map<string, number>();
    let count = 0;
    for (const page of document.getPages()) {
      const annotations = page.node.lookupMaybe(PDFName.of("Annots"), PDFArray);
      if (!annotations) continue;
      count += annotations.size();
      if (count > 20_000) return null;
      for (let position = 0; position < annotations.size(); position++) {
        const reference = annotations.get(position);
        const annotation = document.context.lookup(reference);
        if (!(reference instanceof PDFRef) || !(annotation instanceof PDFDict)) continue;
        const name = document.context.lookup(annotation.get(PDFName.of("NM")));
        if (!(name instanceof PDFString || name instanceof PDFHexString)) continue;
        const text = name.decodeText();
        if (!text.startsWith("MRNote:")) continue;
        claims.set(text, (claims.get(text) ?? 0) + 1);
        byRef.set(`${reference.objectNumber}R${reference.generationNumber || ""}`, text);
      }
    }
    return new Map([...byRef].filter(([, text]) => claims.get(text) === 1));
  } catch { return null; }
}

const annotationNamesByDocument = new WeakMap<PdfDataSource, Promise<ReadonlyMap<string, string> | null>>();

/** One bounded read per pdf.js document incarnation, like the note index. */
export function managedPdfAnnotationNamesForDocument(
  document: PdfDataSource,
): Promise<ReadonlyMap<string, string> | null> {
  let pending = annotationNamesByDocument.get(document);
  if (!pending) {
    pending = document.getData().then(readManagedPdfAnnotationNames).catch(() => null);
    annotationNamesByDocument.set(document, pending);
  }
  return pending;
}
