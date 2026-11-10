/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  PDFArray,
  PDFDict,
  PDFHexString,
  PDFName,
  PDFNumber,
  PDFRef,
  PDFString,
  type PDFObject,
} from "pdf-lib";

import type { EditableAnnotationSpec } from "@mrnote/core/web/lib/pdf/editable-annotations.ts";
import { loadPdfForInspection } from "./encrypted-guard";
import { nativeNoteIdToken } from "@mrnote/core/web/lib/pdf/native-note-roundtrip.ts";
import {
  normalizeXfdfPdfDate,
  normalizeXfdfSourceId,
  XFDF_MAX_ANNOTATIONS,
  type XfdfAnnotationReply,
  type XfdfAnnotationSpec,
  type XfdfAnnotationSubtype,
} from "@mrnote/core/web/lib/pdf/xfdf.ts";















const SUPPORTED_SUBTYPES: ReadonlySet<XfdfAnnotationSubtype> = new Set([
  "FreeText",
  "Highlight",
  "Underline",
  "StrikeOut",
  "Squiggly",
  "Text",
  "Ink",
  "Square",
  "Circle",
  "Line",
  "Polygon",
  "PolyLine",
]);

/** Annotation subtypes that are page furniture rather than comments. */
const NON_COMMENT_SUBTYPES: ReadonlySet<string> = new Set([
  "Link",
  "Widget",
  "Popup",
  "PrinterMark",
  "TrapNet",
  "Watermark",
  "Screen",
  "Movie",
  "3D",
  "RichMedia",
  "Projection",
]);

const MARKUP_SUBTYPES: ReadonlySet<XfdfAnnotationSubtype> = new Set([
  "Highlight",
  "Underline",
  "StrikeOut",
  "Squiggly",
]);

const LINE_ENDINGS: Readonly<
  Record<string, NonNullable<EditableAnnotationSpec["lineStart"]>>
> = {
  OpenArrow: "openArrow",
  ClosedArrow: "closedArrow",
  Circle: "circle",
  Square: "square",
};

const XML_TEXT_UNSAFE = /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/gu;
const MANAGED_NAME_PREFIX = "MRNote:";

export interface ReadNativeXfdfOptions {
  /** Names already exported by the caller (MRAI's own comments, which it
   * also writes into the file with /NM set to their id). */
  readonly excludeNames?: ReadonlySet<string>;
  


  readonly excludeManagedIds?: ReadonlySet<string>;
}

export interface NativeXfdfExport {
  readonly specs: XfdfAnnotationSpec[];
  /** Comments the source carries that the export leaves out, by subtype
   * ("Stamp", "FileAttachment" …). "Reply" counts replies whose parent is not
   * exported; "Limit" counts what fell past the annotation-count bound. */
  readonly omitted: Readonly<Record<string, number>>;
}

export async function readNativeXfdfAnnotationSpecs(
  bytes: Uint8Array,
  options: ReadNativeXfdfOptions = {},
): Promise<XfdfAnnotationSpec[]> {
  return (await readNativeXfdfExport(bytes, options)).specs;
}

export async function readNativeXfdfExport(
  bytes: Uint8Array,
  options: ReadNativeXfdfOptions = {},
): Promise<NativeXfdfExport> {
  const doc = await loadPdfForInspection(bytes, { updateMetadata: false });
  if (!doc) return { specs: [], omitted: {} };
  const exclude = options.excludeNames ?? new Set<string>();
  const excludeManaged = options.excludeManagedIds ?? new Set<string>();
  const specs: XfdfAnnotationSpec[] = [];
  const specByRef = new Map<string, XfdfAnnotationSpec>();
  const excludedRefs = new Set<string>();
  const claimed = new Set<string>(exclude);
  const replies: Array<{ parent: string; dict: PDFDict; pageIndex: number }> =
    [];
  const omitted: Record<string, number> = {};
  const omit = (kind: string) => {
    omitted[kind] = (omitted[kind] ?? 0) + 1;
  };
  let emitted = 0;

  const isCallerExported = (dict: PDFDict): boolean => {
    const own = normalizeXfdfSourceId(text(dict.lookup(PDFName.of("NM"))));
    if (!own) return false;
    if (exclude.has(own)) return true;
    const id = managedIdFromName(own);
    return id !== null && excludeManaged.has(id);
  };

  const claimName = (dict: PDFDict, fallback: string): string => {
    const own = normalizeXfdfSourceId(text(dict.lookup(PDFName.of("NM"))));
    let name = own && !claimed.has(own) ? own : fallback;
    for (let suffix = 2; claimed.has(name); suffix += 1) {
      name = `${fallback}-${suffix}`;
    }
    claimed.add(name);
    return name;
  };

  doc.getPages().forEach((page, pageIndex) => {
    const annots = page.node.lookup(PDFName.of("Annots"));
    if (!(annots instanceof PDFArray)) return;
    for (let index = 0; index < annots.size(); index += 1) {
      const raw = annots.get(index);
      const dict = annots.lookup(index);
      if (!(dict instanceof PDFDict)) continue;
      const subtype = nameOf(dict.lookup(PDFName.of("Subtype")));
      if (!subtype || NON_COMMENT_SUBTYPES.has(subtype)) continue;
      if (isCallerExported(dict)) {
        if (raw instanceof PDFRef) excludedRefs.add(raw.toString());
        continue;
      }
      const inReplyTo = dict.get(PDFName.of("IRT"));
      if (inReplyTo instanceof PDFRef && subtype === "Text") {
        replies.push({ parent: inReplyTo.toString(), dict, pageIndex });
        continue;
      }
      if (!SUPPORTED_SUBTYPES.has(subtype as XfdfAnnotationSubtype)) {
        omit(subtype);
        continue;
      }
      if (emitted >= XFDF_MAX_ANNOTATIONS) {
        omit("Limit");
        continue;
      }
      const fallback = `native-p${pageIndex + 1}-${index + 1}`;
      const name = claimName(dict, fallback);
      const spec = specFromDict(
        dict,
        subtype as XfdfAnnotationSubtype,
        pageIndex,
        name,
      );
      if (!spec) {
        omit(subtype);
        continue;
      }
      specs.push(spec);
      emitted += 1;
      if (raw instanceof PDFRef) specByRef.set(raw.toString(), spec);
    }
  });

  for (const reply of replies) {
    const parent = specByRef.get(reply.parent);
    const contents = cleanText(text(reply.dict.lookup(PDFName.of("Contents"))));
    // An empty reply carries nothing to export; a reply to a comment the
    // caller exports itself travels in that comment's own replies.
    if (!contents || excludedRefs.has(reply.parent)) continue;
    if (!parent) {
      omit("Reply");
      continue;
    }
    if (emitted >= XFDF_MAX_ANNOTATIONS) {
      omit("Limit");
      continue;
    }
    const name = claimName(
      reply.dict,
      `${parent.name ?? "native"}-reply-${(parent.replies?.length ?? 0) + 1}`,
    );
    const entry: XfdfAnnotationReply = {
      name,
      contents,
      author: cleanText(text(reply.dict.lookup(PDFName.of("T")))),
    };
    const modifiedDate = pdfDate(reply.dict, "M");
    const creationDate = pdfDate(reply.dict, "CreationDate");
    if (modifiedDate) entry.modifiedDate = modifiedDate;
    if (creationDate) entry.creationDate = creationDate;
    parent.replies = [...(parent.replies ?? []), entry];
    emitted += 1;
  }
  return { specs, omitted };
}

/** Omission kinds the viewer has a label for; anything else is shown by its
 * PDF subtype name. */
const LABELLED_OMISSION_KINDS: ReadonlySet<string> = new Set([
  "Stamp",
  "FileAttachment",
  "Caret",
  "Sound",
  "Redact",
  "Reply",
  "Limit",
]);

export interface NativeXfdfOmissionSummary {
  readonly total: number;
  readonly entries: ReadonlyArray<{
    readonly kind: string;
    readonly count: number;
    /** Whether `kind` has its own label (`pdfViewer.xfdfOmittedKind.<kind>`). */
    readonly known: boolean;
  }>;
}

/** The omissions of an export, largest first, ready to name to the user. */
export function describeNativeXfdfOmissions(
  omitted: Readonly<Record<string, number>>,
): NativeXfdfOmissionSummary {
  const entries = Object.entries(omitted)
    .filter(([, count]) => Number.isSafeInteger(count) && count > 0)
    .map(([kind, count]) => ({
      kind,
      count,
      known: LABELLED_OMISSION_KINDS.has(kind),
    }))
    .sort((a, b) => b.count - a.count || a.kind.localeCompare(b.kind));
  return {
    total: entries.reduce((sum, entry) => sum + entry.count, 0),
    entries,
  };
}

/** The platform id an MRNote-managed annotation name (`MRNote:<token>`)
 * encodes, when the token is exactly that id's canonical encoding. */
function managedIdFromName(name: string): string | null {
  if (!name.startsWith(MANAGED_NAME_PREFIX)) return null;
  const token = name.slice(MANAGED_NAME_PREFIX.length);
  try {
    const id = decodeURIComponent(token);
    return id && nativeNoteIdToken(id) === token ? id : null;
  } catch {
    return null;
  }
}

function pdfDate(dict: PDFDict, key: string): string | undefined {
  return normalizeXfdfPdfDate(text(dict.lookup(PDFName.of(key))));
}

function specFromDict(
  dict: PDFDict,
  subtype: XfdfAnnotationSubtype,
  pageIndex: number,
  name: string,
): XfdfAnnotationSpec | null {
  const rectValues = numbers(dict.lookup(PDFName.of("Rect")));
  if (rectValues.length !== 4) return null;
  const rect: [number, number, number, number] = [
    Math.min(rectValues[0], rectValues[2]),
    Math.min(rectValues[1], rectValues[3]),
    Math.max(rectValues[0], rectValues[2]),
    Math.max(rectValues[1], rectValues[3]),
  ];
  const ownColour = colour(dict.lookup(PDFName.of("C")));
  const spec: XfdfAnnotationSpec = { pageIndex, subtype, rect, name };
  if (ownColour) spec.color = ownColour;
  // A FreeText without /C has a transparent box; any other kind needs a colour.
  else if (subtype !== "FreeText") {
    spec.color =
      MARKUP_SUBTYPES.has(subtype) || subtype === "Text" ? [1, 1, 0] : [0, 0, 0];
  }
  // The whole text: no per-field cap, so nothing is cut without a word
  
  const contents = cleanText(text(dict.lookup(PDFName.of("Contents"))));
  if (contents) spec.contents = contents;
  const author = cleanText(text(dict.lookup(PDFName.of("T"))));
  if (author) spec.author = author;
  const modifiedDate = pdfDate(dict, "M");
  const creationDate = pdfDate(dict, "CreationDate");
  if (modifiedDate) spec.modifiedDate = modifiedDate;
  if (creationDate) spec.creationDate = creationDate;
  if (subtype === "FreeText") {
    const freeText: NonNullable<XfdfAnnotationSpec["freeText"]> = {};
    const appearance = cleanText(text(dict.lookup(PDFName.of("DA"))));
    const style = cleanText(text(dict.lookup(PDFName.of("DS"))));
    const quadding = numberOf(dict.lookup(PDFName.of("Q")));
    const intent = nameOf(dict.lookup(PDFName.of("IT")));
    const callout = numbers(dict.lookup(PDFName.of("CL")));
    const rotation = numberOf(dict.lookup(PDFName.of("Rotate")));
    if (appearance) freeText.defaultAppearance = appearance;
    if (style) freeText.defaultStyle = style;
    if (quadding === 1) freeText.justification = "centered";
    else if (quadding === 2) freeText.justification = "right";
    if (intent && intent.length <= 64) freeText.intent = intent;
    if (callout.length === 4 || callout.length === 6) freeText.callout = callout;
    if (Object.keys(freeText).length > 0) spec.freeText = freeText;
    if (rotation !== undefined && Number.isFinite(rotation) && rotation !== 0) {
      spec.rotation = rotation;
    }
  }

  if (MARKUP_SUBTYPES.has(subtype)) {
    const quads = numbers(dict.lookup(PDFName.of("QuadPoints")));
    spec.quadPoints =
      quads.length >= 8 && quads.length % 8 === 0
        ? quads
        : [rect[0], rect[3], rect[2], rect[3], rect[0], rect[1], rect[2], rect[1]];
  }
  if (subtype === "Ink") {
    const inkList = dict.lookup(PDFName.of("InkList"));
    const strokes: number[][] = [];
    if (inkList instanceof PDFArray) {
      for (let index = 0; index < inkList.size(); index += 1) {
        const stroke = numbers(inkList.lookup(index));
        if (stroke.length >= 4 && stroke.length % 2 === 0) strokes.push(stroke);
      }
    }
    if (strokes.length === 0) return null;
    spec.inkList = strokes;
  }
  if (subtype === "Line") {
    const line = numbers(dict.lookup(PDFName.of("L")));
    if (line.length !== 4) return null;
    spec.line = [line[0], line[1], line[2], line[3]];
  }
  if (subtype === "Polygon" || subtype === "PolyLine") {
    const vertices = numbers(dict.lookup(PDFName.of("Vertices")));
    const minimum = subtype === "Polygon" ? 6 : 4;
    if (vertices.length < minimum || vertices.length % 2 !== 0) return null;
    spec.vertices = vertices;
  }
  if (subtype === "Line" || subtype === "PolyLine") {
    const endings = dict.lookup(PDFName.of("LE"));
    if (endings instanceof PDFArray && endings.size() === 2) {
      const start = LINE_ENDINGS[nameOf(endings.lookup(0)) ?? ""];
      const end = LINE_ENDINGS[nameOf(endings.lookup(1)) ?? ""];
      if (start) spec.lineStart = start;
      if (end) spec.lineEnd = end;
    }
  }
  if (subtype === "Square" || subtype === "Circle" || subtype === "Polygon") {
    const interior = colour(dict.lookup(PDFName.of("IC")));
    if (interior) spec.interiorColor = interior;
  }
  if (!MARKUP_SUBTYPES.has(subtype) && subtype !== "Text") {
    const border = dict.lookup(PDFName.of("BS"));
    const width =
      border instanceof PDFDict
        ? numberOf(border.lookup(PDFName.of("W")))
        : numbers(dict.lookup(PDFName.of("Border")))[2];
    if (width !== undefined && Number.isFinite(width) && width >= 0) {
      spec.borderWidth = width;
    }
    if (border instanceof PDFDict) {
      const dash = numbers(border.lookup(PDFName.of("D")));
      if (
        dash.length === 2 &&
        dash.every((value) => value > 0 && value <= 1_000)
      ) {
        spec.borderDashArray = [dash[0], dash[1]];
      }
    }
  }
  const opacity = numberOf(dict.lookup(PDFName.of("CA")));
  if (opacity !== undefined && opacity >= 0 && opacity < 1) {
    spec.opacity = opacity;
  }
  return spec;
}

function nameOf(value: PDFObject | undefined): string | undefined {
  return value instanceof PDFName ? value.decodeText() : undefined;
}

function text(value: PDFObject | undefined): string | undefined {
  return value instanceof PDFString || value instanceof PDFHexString
    ? value.decodeText()
    : undefined;
}

function numberOf(value: PDFObject | undefined): number | undefined {
  return value instanceof PDFNumber ? value.asNumber() : undefined;
}

function numbers(value: PDFObject | undefined): number[] {
  if (!(value instanceof PDFArray)) return [];
  const out: number[] = [];
  for (let index = 0; index < value.size(); index += 1) {
    const item = numberOf(value.lookup(index));
    if (item === undefined || !Number.isFinite(item)) return [];
    out.push(item);
  }
  return out;
}

/** /C and /IC: gray, RGB or CMYK, as the RGB the XFDF colour attribute takes. */
function colour(value: PDFObject | undefined): [number, number, number] | null {
  const parts = numbers(value).map((part) => Math.max(0, Math.min(1, part)));
  if (parts.length === 1) return [parts[0], parts[0], parts[0]];
  if (parts.length === 3) return [parts[0], parts[1], parts[2]];
  if (parts.length === 4) {
    const [c, m, y, k] = parts;
    return [(1 - c) * (1 - k), (1 - m) * (1 - k), (1 - y) * (1 - k)];
  }
  return null;
}

function cleanText(value: string | undefined): string | undefined {
  if (!value) return undefined;
  const cleaned = value.replace(XML_TEXT_UNSAFE, "");
  return cleaned.trim() ? cleaned : undefined;
}
