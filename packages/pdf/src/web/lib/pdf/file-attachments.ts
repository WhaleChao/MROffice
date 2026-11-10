/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * File attachments pinned to a PLACE on a page.
 *
 * The document-level attachment list (which formula MathML already uses) says
 * "this file travels with the document". A `/FileAttachment` annotation says
 * something stronger: "this file belongs HERE" — the paperclip next to the
 * paragraph it explains, the spreadsheet pinned to the table it backs. Acrobat
 * and WPS both surface it in place, and it is the difference between an
 * appendix and a footnote.
 *
 * Every attachment is registered twice on purpose: as the annotation at its
 * place, and in the catalog's `/Names /EmbeddedFiles` tree. A reader that
 * ignores annotations still lists the file in its attachments panel, so the
 * bytes are never reachable only through a feature the reader lacks.
 */

import {
  PDFArray,
  PDFDict,
  PDFHexString,
  PDFName,
  PDFNull,
  PDFNumber,
  PDFRawStream,
  PDFRef,
  PDFString,
  type PDFDocument,
  type PDFObject,
} from "pdf-lib";
import { Unzlib } from "fflate";

import { PDF_EDITABLE_MAX_BYTES } from "@mrnote/core/web/lib/document-size-limits.ts";

import {
  isSafePdfFileAttachmentDownloadName,
} from "./native-file-attachment-roundtrip";

/** The icons ISO 32000-1 Table 184 names. */
export type FileAttachmentIcon = "Paperclip" | "PushPin" | "Graph" | "Tag";

export const FILE_ATTACHMENT_ICONS: readonly FileAttachmentIcon[] = [
  "Paperclip",
  "PushPin",
  "Graph",
  "Tag",
];

/** Drawn icon size in points — the box the annotation occupies. */
export const FILE_ATTACHMENT_ICON_SIZE = 18;

/** No attachment can outgrow the PDF it is in: the editable PDF's document
 * total (lib/document-size-limits) bounds one, as it bounds them all. */
export const FILE_ATTACHMENT_MAX_BYTES = PDF_EDITABLE_MAX_BYTES;

export interface FileAttachmentRecord {
  id: string;
  /** 1-based. */
  page: number;
  /** Anchor point in page-relative percent [0..100], top-left origin. */
  x: number;
  y: number;
  /** File name shown to the reader; sanitised before writing. */
  name: string;
  bytes: Uint8Array;
  mimeType: string;
  /** Shown as the annotation's comment. */
  description?: string;
  author?: string;
  /** PDF date string, supplied rather than read from the clock. */
  created?: string;
  icon?: FileAttachmentIcon;
  /** `#rrggbb`. */
  color?: string;
}

export interface WriteFileAttachmentsResult {
  written: number;
  skipped: Array<{ id: string; reason: string }>;
}

/** An attachment read back off a page; unsafe bytes remain quarantined. */
export interface ReadFileAttachment {
  page: number;
  /** Percent [0..100], top-left origin — the anchor's centre. */
  x: number;
  y: number;
  name: string;
  description: string;
  /** The annotation's /NM: the identity MRAI's own attachments carry. */
  annotationName?: string;
  bytes: Uint8Array;
  size: number;
  icon: FileAttachmentIcon;
  blockedReason?: FileAttachmentReadBlockedReason;
}

export type FileAttachmentReadBlockedReason =
  | "dangerous_name"
  | "empty_file"
  | "too_large"
  | "unsupported_filter"
  | "size_mismatch"
  | "decode_failed"
  | "missing_stream";

export type BoundedFileAttachmentDecode =
  // `blockedReason?: undefined` keeps the success branch discriminable, so a
  // caller can read `decoded.blockedReason` across every decode outcome.
  | { bytes: Uint8Array; size: number; blockedReason?: undefined }
  | {
      bytes: Uint8Array;
      size: number;
      blockedReason: Exclude<
        FileAttachmentReadBlockedReason,
        "dangerous_name" | "missing_stream"
      >;
    };

function decodeText(value: unknown): string {
  if (value instanceof PDFHexString || value instanceof PDFString) {
    return value.decodeText();
  }
  return "";
}

function declaredAttachmentSize(stream: PDFRawStream): number | null | undefined {
  const params = stream.dict.lookupMaybe(PDFName.of("Params"), PDFDict);
  if (!params) return undefined;
  const raw = params.lookupMaybe(PDFName.of("Size"), PDFNumber);
  if (!raw) return undefined;
  const size = raw.asNumber();
  return Number.isSafeInteger(size) && size >= 0 ? size : null;
}








function attachmentFilterChain(stream: PDFRawStream): string[] | null {
  const filterKey = PDFName.of("Filter");
  if (stream.dict.get(filterKey) === undefined) return [];
  const filter = stream.dict.lookup(filterKey);
  if (filter === undefined || filter === PDFNull) return [];
  if (filter instanceof PDFName) return [filter.decodeText()];
  if (!(filter instanceof PDFArray)) return null;
  const chain: string[] = [];
  for (let index = 0; index < filter.size(); index += 1) {
    const entry = filter.lookup(index);
    if (!(entry instanceof PDFName)) return null;
    chain.push(entry.decodeText());
  }
  return chain;
}

/** The one filter's /DecodeParms, resolved (a one-element array unwrapped). */
function attachmentSingleDecodeParms(
  stream: PDFRawStream,
): PDFDict | "none" | "invalid" {
  const key = PDFName.of("DecodeParms");
  if (stream.dict.get(key) === undefined) return "none";
  let params = stream.dict.lookup(key);
  if (params instanceof PDFArray) {
    if (params.size() !== 1) return "invalid";
    params = params.lookup(0);
  }
  if (params === undefined || params === PDFNull) return "none";
  return params instanceof PDFDict ? params : "invalid";
}

/**
 * Only an unfiltered stream or exactly one Flate filter (without a predictor)
 * is decoded; any other or longer chain stays quarantined.
 */
function attachmentFlateFilterIsSupported(stream: PDFRawStream): boolean {
  const chain = attachmentFilterChain(stream);
  if (chain === null) return false;
  if (chain.length === 0) return true;
  if (chain.length !== 1) return false;
  if (chain[0] !== "FlateDecode" && chain[0] !== "Fl") return false;
  const params = attachmentSingleDecodeParms(stream);
  if (params === "none") return true;
  if (params === "invalid") return false;
  const keys = params.keys().map(String);
  if (keys.some((key) => key !== "/Predictor")) return false;
  const predictor = params.lookupMaybe(PDFName.of("Predictor"), PDFNumber);
  return predictor === undefined || predictor.asNumber() === 1;
}

function joinAttachmentChunks(
  chunks: readonly Uint8Array[],
  size: number,
): Uint8Array {
  const output = new Uint8Array(size);
  let offset = 0;
  for (const chunk of chunks) {
    output.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return output;
}

/**
 * Decode an imported attachment without ever asking a decompressor for an
 * unbounded result. Tiny compressed input chunks keep each synchronous fflate
 * output bounded; the callback aborts before the aggregate crosses the hard
 * limit. Unsupported filters and predictor pipelines remain visible but
 * quarantined instead of being passed to pdf-lib's unbounded decoder.
 */
export function decodeFileAttachmentStreamBounded(
  stream: PDFRawStream,
  maxBytes = FILE_ATTACHMENT_MAX_BYTES,
): BoundedFileAttachmentDecode {
  const declared = declaredAttachmentSize(stream);
  if (
    !Number.isSafeInteger(maxBytes) ||
    maxBytes < 1 ||
    declared === null
  ) {
    return { bytes: new Uint8Array(0), size: 0, blockedReason: "size_mismatch" };
  }
  if (declared !== undefined && declared > maxBytes) {
    return { bytes: new Uint8Array(0), size: declared, blockedReason: "too_large" };
  }
  const source = stream.getContents();
  const chain = attachmentFilterChain(stream);
  if (chain !== null && chain.length === 0) {
    if (source.byteLength > maxBytes) {
      return {
        bytes: new Uint8Array(0),
        size: declared ?? source.byteLength,
        blockedReason: "too_large",
      };
    }
    if (source.byteLength === 0) {
      return { bytes: new Uint8Array(0), size: 0, blockedReason: "empty_file" };
    }
    if (declared !== undefined && declared !== source.byteLength) {
      return {
        bytes: new Uint8Array(0),
        size: declared,
        blockedReason: "size_mismatch",
      };
    }
    return { bytes: source.slice(), size: source.byteLength };
  }
  if (!attachmentFlateFilterIsSupported(stream)) {
    return {
      bytes: new Uint8Array(0),
      size: declared ?? source.byteLength,
      blockedReason: "unsupported_filter",
    };
  }

  const chunks: Uint8Array[] = [];
  let total = 0;
  let complete = false;
  const tooLarge = new Error("attachment-too-large");
  try {
    const decoder = new Unzlib((chunk, final) => {
      if (total + chunk.byteLength > maxBytes) throw tooLarge;
      chunks.push(chunk.slice());
      total += chunk.byteLength;
      complete = final;
    });
    const inputChunkBytes = 1024;
    for (let offset = 0; offset < source.byteLength; offset += inputChunkBytes) {
      const end = Math.min(offset + inputChunkBytes, source.byteLength);
      decoder.push(source.subarray(offset, end), end === source.byteLength);
    }
  } catch (error) {
    return {
      bytes: new Uint8Array(0),
      size: declared ?? total,
      blockedReason: error === tooLarge ? "too_large" : "decode_failed",
    };
  }
  if (!complete || total === 0) {
    return {
      bytes: new Uint8Array(0),
      size: declared ?? total,
      blockedReason: total === 0 ? "empty_file" : "decode_failed",
    };
  }
  if (declared !== undefined && declared !== total) {
    return {
      bytes: new Uint8Array(0),
      size: declared,
      blockedReason: "size_mismatch",
    };
  }
  return { bytes: joinAttachmentChunks(chunks, total), size: total };
}

/**
 * Read every `/FileAttachment` on every page. Safe bounded bytes are included;
 * quarantined objects retain their page/name metadata and an explicit reason.
 *
 * This is the import half: a file a colleague pinned in Acrobat shows up with
 * its place, its name and its content, not as an unexplained icon.
 */
export function readFileAttachmentAnnotations(
  doc: PDFDocument,
  pageSizes: ReadonlyArray<{ width: number; height: number }>,
): ReadFileAttachment[] {
  const out: ReadFileAttachment[] = [];
  const pages = doc.getPages();
  pages.forEach((page, index) => {
    const size = pageSizes[index];
    if (!size || !(size.width > 0) || !(size.height > 0)) return;
    const annots = page.node.lookupMaybe(PDFName.of("Annots"), PDFArray);
    if (!annots) return;
    for (let i = 0; i < annots.size(); i += 1) {
      const annotation = doc.context.lookupMaybe(annots.get(i), PDFDict);
      if (!annotation) continue;
      if (String(annotation.get(PDFName.of("Subtype"))) !== "/FileAttachment") {
        continue;
      }
      const spec = annotation.lookupMaybe(PDFName.of("FS"), PDFDict);
      const ef = spec?.lookupMaybe(PDFName.of("EF"), PDFDict);
      // /EF keys mirror the file-name keys: a PDF 1.7 file may embed the
      // stream under /UF (Unicode name) alone, which is as legitimate as /F
      
      const streamRef =
        ef?.get(PDFName.of("UF")) ?? ef?.get(PDFName.of("F"));
      const stream = streamRef ? doc.context.lookup(streamRef) : null;
      // The names may be indirect objects: read through the reference, or a
      // legitimate "notes.txt" became the placeholder "attachment", which has
      // no extension and was then quarantined as possibly executable
      
      const name =
        decodeText(spec?.lookup(PDFName.of("UF"))) ||
        decodeText(spec?.lookup(PDFName.of("F"))) ||
        "attachment";
      const decoded =
        !isSafePdfFileAttachmentDownloadName(name)
          ? {
              bytes: new Uint8Array(0),
              size:
                stream instanceof PDFRawStream
                  ? (declaredAttachmentSize(stream) ?? stream.getContentsSize())
                  : 0,
              blockedReason: "dangerous_name" as const,
            }
          : stream instanceof PDFRawStream
            ? decodeFileAttachmentStreamBounded(stream)
            : {
                bytes: new Uint8Array(0),
                size: 0,
                blockedReason: "missing_stream" as const,
              };
      const rect = annotation.lookupMaybe(PDFName.of("Rect"), PDFArray);
      let cx = 0;
      let cy = 0;
      if (rect && rect.size() === 4) {
        const numbers = Array.from({ length: 4 }, (_, j) =>
          Number(String(rect.get(j))),
        );
        cx = (numbers[0] + numbers[2]) / 2;
        cy = (numbers[1] + numbers[3]) / 2;
      }
      const iconRaw = String(
        annotation.get(PDFName.of("Name")) ?? "/Paperclip",
      ).replace(/^\//, "");
      out.push({
        page: index + 1,
        x: (cx / size.width) * 100,
        y: ((size.height - cy) / size.height) * 100,
        name,
        description: decodeText(annotation.get(PDFName.of("Contents"))),
        ...(decodeText(annotation.get(PDFName.of("NM")))
          ? { annotationName: decodeText(annotation.get(PDFName.of("NM"))) }
          : {}),
        bytes: decoded.bytes,
        size: decoded.size,
        icon: (FILE_ATTACHMENT_ICONS as readonly string[]).includes(iconRaw)
          ? (iconRaw as FileAttachmentIcon)
          : "Paperclip",
        ...(decoded.blockedReason
          ? { blockedReason: decoded.blockedReason }
          : {}),
      });
    }
  });
  return out;
}

/** A file the document carries in its catalog's `/EmbeddedFiles` tree. */
export interface ReadDocumentAttachment {
  /** The name tree's key for it (unique within the tree). */
  key: string;
  name: string;
  description: string;
  bytes: Uint8Array;
  size: number;
  blockedReason?: FileAttachmentReadBlockedReason;
}

/** The deepest name tree a reader walks, and how many entries it reads. */
const NAME_TREE_MAX_DEPTH = 32;
const NAME_TREE_MAX_ENTRIES = 4_096;

function refKey(value: PDFObject | undefined): string | null {
  return value instanceof PDFRef ? value.toString() : null;
}

/**
 * The file specifications and embedded streams the pages' `/FileAttachment`
 * annotations already list: the same file registered in the name tree too
 * is one file, shown once at its place. Identity is the object, never the
 * content -- two files with the same bytes are two files.
 */
function pageAttachmentObjectKeys(doc: PDFDocument): Set<string> {
  const keys = new Set<string>();
  for (const page of doc.getPages()) {
    try {
      const annots = page.node.lookupMaybe(PDFName.of("Annots"), PDFArray);
      if (!annots) continue;
      for (let i = 0; i < annots.size(); i += 1) {
        try {
          const annotation = doc.context.lookupMaybe(annots.get(i), PDFDict);
          if (String(annotation?.get(PDFName.of("Subtype"))) !== "/FileAttachment") continue;
          const specRef = refKey(annotation?.get(PDFName.of("FS")));
          if (specRef) keys.add(specRef);
          const spec = annotation?.lookupMaybe(PDFName.of("FS"), PDFDict);
          const ef = spec?.lookupMaybe(PDFName.of("EF"), PDFDict);
          for (const key of ["UF", "F"]) {
            const streamRef = refKey(ef?.get(PDFName.of(key)));
            if (streamRef) keys.add(`stream:${streamRef}`);
          }
        } catch {
          // An annotation of the wrong shape lists nothing.
        }
      }
    } catch {
      // A page whose /Annots is not an array lists nothing.
    }
  }
  return keys;
}









export function readDocumentEmbeddedFiles(doc: PDFDocument): ReadDocumentAttachment[] {
  let root: PDFDict | undefined;
  try {
    root = doc.catalog
      .lookupMaybe(PDFName.of("Names"), PDFDict)
      ?.lookupMaybe(PDFName.of("EmbeddedFiles"), PDFDict);
  } catch {
    root = undefined;
  }
  if (!root) return [];
  const listed = pageAttachmentObjectKeys(doc);
  const out: ReadDocumentAttachment[] = [];
  const seenNodes = new Set<PDFDict>();
  const seenSpecs = new Set<string>();
  let entries = 0;
  const visit = (node: PDFDict, depth: number) => {
    if (depth > NAME_TREE_MAX_DEPTH || seenNodes.has(node)) return;
    seenNodes.add(node);
    let pairs: PDFArray | undefined;
    try {
      pairs = node.lookupMaybe(PDFName.of("Names"), PDFArray);
    } catch {
      pairs = undefined;
    }
    if (pairs) {
      for (let i = 0; i + 1 < pairs.size(); i += 2) {
        if (entries >= NAME_TREE_MAX_ENTRIES) return;
        entries += 1;
        try {
          readEntry(pairs, i);
        } catch {
          // pdf-lib throws on a value of the wrong type: that entry is
          // unreadable, the rest of the tree is not.
        }
      }
    }
    let kids: PDFArray | undefined;
    try {
      kids = node.lookupMaybe(PDFName.of("Kids"), PDFArray);
    } catch {
      kids = undefined;
    }
    if (!kids) return;
    for (let i = 0; i < kids.size(); i += 1) {
      let kid: PDFDict | undefined;
      try {
        kid = doc.context.lookupMaybe(kids.get(i), PDFDict);
      } catch {
        kid = undefined;
      }
      if (kid) visit(kid, depth + 1);
    }
  };
  const readEntry = (pairs: PDFArray, i: number) => {
    const key = decodeText(doc.context.lookup(pairs.get(i)));
    const rawSpec = pairs.get(i + 1);
    const specRef = refKey(rawSpec);
    if (specRef && (listed.has(specRef) || seenSpecs.has(specRef))) return;
    if (specRef) seenSpecs.add(specRef);
    const spec = doc.context.lookupMaybe(rawSpec, PDFDict);
    if (!spec) return;
    const ef = spec.lookupMaybe(PDFName.of("EF"), PDFDict);
    const streamEntry = ef?.get(PDFName.of("UF")) ?? ef?.get(PDFName.of("F"));
    const streamRef = refKey(streamEntry);
    if (streamRef && listed.has(`stream:${streamRef}`)) return;
    const stream = streamEntry ? doc.context.lookup(streamEntry) : null;
    const name =
      decodeText(spec.lookup(PDFName.of("UF"))) ||
      decodeText(spec.lookup(PDFName.of("F"))) ||
      key ||
      "attachment";
    const decoded = !isSafePdfFileAttachmentDownloadName(name)
      ? {
          bytes: new Uint8Array(0),
          size:
            stream instanceof PDFRawStream
              ? (declaredAttachmentSize(stream) ?? stream.getContentsSize())
              : 0,
          blockedReason: "dangerous_name" as const,
        }
      : stream instanceof PDFRawStream
        ? decodeFileAttachmentStreamBounded(stream)
        : {
            bytes: new Uint8Array(0),
            size: 0,
            blockedReason: "missing_stream" as const,
          };
    out.push({
      key,
      name,
      description: decodeText(spec.lookup(PDFName.of("Desc"))),
      bytes: decoded.bytes,
      size: decoded.size,
      ...(decoded.blockedReason ? { blockedReason: decoded.blockedReason } : {}),
    });
  };
  try {
    visit(root, 0);
  } catch {
    // A malformed tree lists what was read before it broke.
  }
  return out;
}
