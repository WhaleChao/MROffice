/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * XFDF — the interchange half of editable annotations.
 *
 * Writing real /Annots makes an exported PDF editable in Acrobat and WPS. XFDF
 * is the other direction: a small XML file carrying only the annotations, which
 * is how a supervisor sends back "here are my comments" without shipping the
 * whole document, and how Acrobat's own Comments > Export/Import works.
 *
 * Both directions key on `name` — our annotation id, written as /NM. That is
 * what makes a merge IDEMPOTENT: importing the same file twice updates the same
 * annotations instead of stacking a second copy on top of the first, which is
 * the failure everyone who has round-tripped comments has seen.
 *
 * Deliberately hand-rolled rather than pulled from a library: the subset we need
 * is a dozen elements, and the escaping and the 0-based page numbering are
 * exactly the details a generic serializer gets wrong for us.
 */

import type {
  EditableAnnotationReply,
  EditableAnnotationSpec,
  EditableAnnotationSubtype,
} from "./editable-annotations";

/**
 * What an XFDF export can carry: every editable subtype, plus FreeText, which
 * is only ever read out of the source PDF (the editor has no FreeText model)
 * and written through unchanged so an export does not silently drop it.
 */
export type XfdfAnnotationSubtype = EditableAnnotationSubtype | "FreeText";

/** XFDF element name for each subtype we emit. */
const ELEMENT_FOR: Record<XfdfAnnotationSubtype, string> = {
  FreeText: "freetext",
  Highlight: "highlight",
  Underline: "underline",
  StrikeOut: "strikeout",
  Squiggly: "squiggly",
  Text: "text",
  Ink: "ink",
  Square: "square",
  Circle: "circle",
  Line: "line",
  Polygon: "polygon",
  PolyLine: "polyline",
};

const SUBTYPE_FOR: Record<string, EditableAnnotationSubtype> = {
  highlight: "Highlight",
  underline: "Underline",
  strikeout: "StrikeOut",
  squiggly: "Squiggly",
  text: "Text",
  ink: "Ink",
  square: "Square",
  circle: "Circle",
  line: "Line",
  polygon: "Polygon",
  polyline: "PolyLine",
};

const XFDF_LINE_ENDING = {
  openArrow: "OpenArrow",
  closedArrow: "ClosedArrow",
  circle: "Circle",
  square: "Square",
} as const;

export const XFDF_SOURCE_ID_MAX_LENGTH = 240;
export const XFDF_MAX_CODE_UNITS = 2_000_000;
export const XFDF_MAX_FILE_BYTES = 4_000_000;
export const XFDF_MAX_ANNOTATIONS = 5_000;
const XFDF_MAX_CONTENT_CODE_UNITS = 65_536;
const XFDF_MAX_AUTHOR_CODE_UNITS = 1_024;
const XFDF_MAX_HREF_CODE_UNITS = 2_048;
const XFDF_MAX_FREETEXT_STYLE = 4_096;
const XFDF_MAX_INK_GESTURES = 500;
const XFDF_MAX_INK_VALUES_PER_GESTURE = 20_000;
const XFDF_NAMESPACE = "http://ns.adobe.com/xfdf/";
const MRAI_XFDF_NAMESPACE = "https://mrnote.com/ns/pdf-xfdf/1";
const XFDF_SOURCE_ID_UNSAFE = /[\u0000-\u001f\u007f]/u;
const XFDF_XML_TEXT_UNSAFE = /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/u;
const SHA256_PATTERN = /^[0-9a-f]{64}$/;

/** A bounded, inert XFDF annotation identity (`name` / PDF `/NM`). */
export function normalizeXfdfSourceId(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const normalized = value.normalize("NFC").trim();
  if (
    !normalized ||
    normalized.length > XFDF_SOURCE_ID_MAX_LENGTH ||
    XFDF_SOURCE_ID_UNSAFE.test(normalized)
  ) {
    return undefined;
  }
  return normalized;
}

/** XML text/attribute escaping. Apostrophe included because attributes are
 *  written with double quotes but a consumer may re-serialize with single. */
export function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function hexColor(color: readonly [number, number, number]): string {
  const part = (v: number) => {
    const clamped = Math.max(0, Math.min(1, Number.isFinite(v) ? v : 0));
    return Math.round(clamped * 255)
      .toString(16)
      .padStart(2, "0")
      .toUpperCase();
  };
  return `#${part(color[0])}${part(color[1])}${part(color[2])}`;
}

function parseHexColor(value: string | null): [number, number, number] | null {
  if (!value) return null;
  const match = /^#?([0-9a-f]{6})$/i.exec(value.trim());
  if (!match) return null;
  const n = Number.parseInt(match[1], 16);
  return [((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255];
}

function numberList(value: string | null, maxValues = 100_000): number[] {
  if (!value) return [];
  const parts = value.split(/[,;\s]+/).filter(Boolean);
  if (parts.length > maxValues) return [];
  const numbers = parts.map((part) => Number(part));
  return numbers.every((number) => Number.isFinite(number)) ? numbers : [];
}

/** FreeText entries XFDF carries as their own attributes and elements. */
export interface XfdfFreeTextFields {
  /** /DA — `<defaultappearance>`, the font, size and colour of the text. */
  defaultAppearance?: string;
  /** /DS — `<defaultstyle>`, the CSS-like default style string. */
  defaultStyle?: string;
  /** /Q — `justification`. */
  justification?: "left" | "centered" | "right";
  /** /IT — `intent` (FreeTextCallout, FreeTextTypeWriter). */
  intent?: string;
  /** /CL — `callout`, 4 or 6 numbers. */
  callout?: number[];
}

/** What XFDF carries of an annotation: everything but its drawn appearance,
 * which a reader regenerates from these entries. */
export type XfdfAnnotationSpec = Omit<
  EditableAnnotationSpec,
  "appearance" | "subtype" | "color" | "replies"
> & {
  subtype: XfdfAnnotationSubtype;
  /** /C. Optional only for FreeText, where an absent /C means a transparent
   * box: inventing a colour would paint a background the source never had. */
  color?: [number, number, number];
  /** /M, as the PDF date string the source carries (XFDF `date`). */
  modifiedDate?: string;
  /** /CreationDate, as the PDF date string (XFDF `creationdate`). */
  creationDate?: string;
  /** /Rotate of a FreeText (XFDF `rotation`), in degrees. */
  rotation?: number;
  freeText?: XfdfFreeTextFields;
  replies?: XfdfAnnotationReply[];
};

export interface XfdfAnnotationReply extends EditableAnnotationReply {
  modifiedDate?: string;
  creationDate?: string;
}

/** A PDF date string (`D:YYYYMMDDHHmmSSOHH'mm'`, every part after the year
 * optional) as XFDF's `date` / `creationdate` attributes carry it. */
const PDF_DATE_STRING =
  /^D:\d{4}(?:\d{2}(?:\d{2}(?:\d{2}(?:\d{2}(?:\d{2})?)?)?)?)?(?:Z(?:00'?(?:00'?)?)?|[+-]\d{2}(?:'?\d{2}'?)?)?$/;

/** The source's own date when it is a well-formed PDF date, else undefined. */
export function normalizeXfdfPdfDate(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  const prefixed = trimmed.startsWith("D:") ? trimmed : `D:${trimmed}`;
  return prefixed.length <= 64 && PDF_DATE_STRING.test(prefixed)
    ? prefixed
    : undefined;
}

export interface ToXfdfOptions {
  /** `<f href>` — the document these annotations belong to. */
  href?: string;
  /** Timestamp for annotations that carry no date of their own. */
  date?: string;
  /** Exact source bytes for MRAI-to-MRAI cross-document refusal. */
  sourceSha256?: string;
}

/** Serialize annotation specs to an XFDF document. */
export function toXfdf(
  specs: ReadonlyArray<XfdfAnnotationSpec>,
  options: ToXfdfOptions = {},
): string {
  assertXfdfExportInput(specs, options);
  const sourceSha256 = options.sourceSha256?.trim().toLowerCase();
  const exactSourceSha256 =
    sourceSha256 && SHA256_PATTERN.test(sourceSha256)
      ? sourceSha256
      : undefined;
  const hasShapeExtensions = specs.some(
    (spec) =>
      (spec.shapeFillOpacity !== undefined && spec.shapeFillOpacity !== 1) ||
      (spec.shapeCornerRadius !== undefined && spec.shapeCornerRadius !== 0),
  );
  const lines: string[] = [
    '<?xml version="1.0" encoding="UTF-8"?>',
    `<xfdf xmlns="${XFDF_NAMESPACE}"${
      exactSourceSha256 || hasShapeExtensions
        ? ` xmlns:mrai="${MRAI_XFDF_NAMESPACE}"`
        : ""
    } xml:space="preserve">`,
  ];
  if (options.href) {
    lines.push(`  <f href="${escapeXml(options.href)}"/>`);
  }
  if (exactSourceSha256) {
    lines.push(`  <mrai:source sha256="${exactSourceSha256}"/>`);
  }
  lines.push("  <annots>");
  for (const spec of specs) {
    lines.push(...annotElement(spec, options));
  }
  lines.push("  </annots>");
  lines.push("</xfdf>");
  const xml = lines.join("\n");
  if (xml.length > XFDF_MAX_CODE_UNITS) {
    throw new Error("XFDF output exceeds the bounded file size");
  }
  return xml;
}

function assertXfdfExportInput(
  specs: ReadonlyArray<XfdfAnnotationSpec>,
  options: ToXfdfOptions,
): void {
  if (specs.length > XFDF_MAX_ANNOTATIONS) {
    throw new Error("XFDF output exceeds the bounded annotation count");
  }
  if (
    (options.href !== undefined &&
      (!options.href ||
        options.href.length > XFDF_MAX_HREF_CODE_UNITS ||
        XFDF_XML_TEXT_UNSAFE.test(options.href))) ||
    (options.date !== undefined &&
      (!options.date ||
        options.date.length > 128 ||
        XFDF_XML_TEXT_UNSAFE.test(options.date)))
  ) {
    throw new Error("XFDF output metadata is invalid");
  }
  const seenNames = new Set<string>();
  let emittedObjects = 0;
  const claimName = (value: string | undefined) => {
    if (value === undefined) return;
    const normalized = normalizeXfdfSourceId(value);
    if (!normalized || normalized !== value || seenNames.has(normalized)) {
      throw new Error("XFDF output annotation identity is invalid");
    }
    seenNames.add(normalized);
  };
  const validText = (value: string | undefined, maxLength: number) =>
    value === undefined ||
    (value.length <= maxLength && !XFDF_XML_TEXT_UNSAFE.test(value));
  const validDate = (value: string | undefined) =>
    value === undefined || normalizeXfdfPdfDate(value) === value;
  const validNumbers = (
    values: readonly number[] | undefined,
    minimum: number,
    multiple: number,
  ) =>
    values === undefined ||
    (values.length >= minimum &&
      values.length % multiple === 0 &&
      values.every((value) => Number.isFinite(value)));

  for (const spec of specs) {
    emittedObjects += 1;
    claimName(spec.name);
    if (
      !Number.isSafeInteger(spec.pageIndex) ||
      spec.pageIndex < 0 ||
      spec.rect.length !== 4 ||
      !spec.rect.every((value) => Number.isFinite(value)) ||
      !(spec.subtype in ELEMENT_FOR) ||
      (spec.color === undefined
        ? spec.subtype !== "FreeText"
        : spec.color.length !== 3 ||
          !spec.color.every(
            (value) => Number.isFinite(value) && value >= 0 && value <= 1,
          )) ||
      // Comment text is written whole: the file-size bound below is the only
      // limit, and exceeding it fails the export loudly. A per-field cap here
      
      !validText(spec.contents, XFDF_MAX_CODE_UNITS) ||
      !validText(spec.author, XFDF_MAX_CODE_UNITS) ||
      !validDate(spec.modifiedDate) ||
      !validDate(spec.creationDate) ||
      (spec.rotation !== undefined && !Number.isFinite(spec.rotation)) ||
      (spec.freeText !== undefined &&
        (spec.subtype !== "FreeText" ||
          !validText(spec.freeText.defaultAppearance, XFDF_MAX_FREETEXT_STYLE) ||
          !validText(spec.freeText.defaultStyle, XFDF_MAX_FREETEXT_STYLE) ||
          !validText(spec.freeText.intent, 64) ||
          (spec.freeText.justification !== undefined &&
            !["left", "centered", "right"].includes(
              spec.freeText.justification,
            )) ||
          (spec.freeText.callout !== undefined &&
            ((spec.freeText.callout.length !== 4 &&
              spec.freeText.callout.length !== 6) ||
              !spec.freeText.callout.every((value) =>
                Number.isFinite(value),
              ))))) ||
      !validNumbers(spec.quadPoints, 8, 8) ||
      !validNumbers(spec.line, 4, 4) ||
      ((spec.subtype === "Polygon" || spec.subtype === "PolyLine") &&
        spec.vertices === undefined) ||
      !validNumbers(spec.vertices, spec.subtype === "Polygon" ? 6 : 4, 2) ||
      !validNumbers(spec.interiorColor, 3, 3) ||
      (spec.borderDashArray !== undefined &&
        (!validNumbers(spec.borderDashArray, 2, 2) ||
          spec.borderDashArray.some((value) => value <= 0 || value > 1_000))) ||
      [spec.lineStart, spec.lineEnd].some(
        (value) => value !== undefined && !(value in XFDF_LINE_ENDING),
      ) ||
      ((spec.lineStart !== undefined || spec.lineEnd !== undefined) &&
        spec.subtype !== "Line" &&
        spec.subtype !== "PolyLine") ||
      (spec.borderWidth !== undefined &&
        (!Number.isFinite(spec.borderWidth) || spec.borderWidth < 0)) ||
      (spec.opacity !== undefined &&
        (!Number.isFinite(spec.opacity) ||
          spec.opacity < 0 ||
          spec.opacity > 1)) ||
      (spec.shapeFillOpacity !== undefined &&
        (!Number.isFinite(spec.shapeFillOpacity) ||
          spec.shapeFillOpacity < 0 ||
          spec.shapeFillOpacity > 1 ||
          spec.interiorColor === undefined ||
          (spec.subtype !== "Square" && spec.subtype !== "Circle"))) ||
      (spec.shapeCornerRadius !== undefined &&
        (!Number.isFinite(spec.shapeCornerRadius) ||
          spec.shapeCornerRadius <= 0 ||
          spec.shapeCornerRadius > 50 ||
          spec.subtype !== "Square"))
    ) {
      throw new Error("XFDF output annotation is invalid");
    }
    for (const stroke of spec.inkList ?? []) {
      if (!validNumbers(stroke, 4, 2)) {
        throw new Error("XFDF output ink gesture is invalid");
      }
    }
    for (const reply of spec.replies ?? []) {
      if (!reply.contents) continue;
      emittedObjects += 1;
      claimName(reply.name);
      if (
        !validText(reply.contents, XFDF_MAX_CODE_UNITS) ||
        !validText(reply.author, XFDF_MAX_CODE_UNITS) ||
        !validDate(reply.modifiedDate) ||
        !validDate(reply.creationDate)
      ) {
        throw new Error("XFDF output reply is invalid");
      }
    }
  }
  if (emittedObjects > XFDF_MAX_ANNOTATIONS) {
    throw new Error("XFDF output exceeds the bounded annotation count");
  }
}

function attrs(pairs: Array<[string, string | undefined]>): string {
  return pairs
    .filter((pair): pair is [string, string] => pair[1] !== undefined)
    .map(([key, value]) => `${key}="${escapeXml(value)}"`)
    .join(" ");
}

function annotElement(
  spec: XfdfAnnotationSpec,
  options: ToXfdfOptions,
): string[] {
  const element = ELEMENT_FOR[spec.subtype];
  const common: Array<[string, string | undefined]> = [
    // XFDF pages are 0-based, matching our spec's own pageIndex. Emitting a
    // 1-based number here is the classic off-by-one that lands every comment
    // one page late in Acrobat.
    ["page", String(spec.pageIndex)],
    ["rect", spec.rect.map((n) => round(n)).join(",")],
    ["color", spec.color ? hexColor(spec.color) : undefined],
    ["name", spec.name],
    ["title", spec.author],
    
    // the export time only stands in for an annotation that has none.
    ["date", spec.modifiedDate ?? options.date],
    ["creationdate", spec.creationDate],
  ];
  if (spec.rotation !== undefined && spec.rotation !== 0) {
    common.push(["rotation", round(spec.rotation)]);
  }
  if (spec.freeText?.justification) {
    common.push(["justification", spec.freeText.justification]);
  }
  if (spec.freeText?.intent) {
    common.push(["intent", spec.freeText.intent]);
  }
  if (spec.freeText?.callout) {
    common.push(["callout", spec.freeText.callout.map(round).join(",")]);
  }
  if (spec.quadPoints && spec.quadPoints.length > 0) {
    common.push(["coords", spec.quadPoints.map((n) => round(n)).join(",")]);
  }
  if (spec.line) {
    common.push(["start", `${round(spec.line[0])},${round(spec.line[1])}`]);
    common.push(["end", `${round(spec.line[2])},${round(spec.line[3])}`]);
  }
  if (spec.borderDashArray) {
    common.push(["dashes", spec.borderDashArray.map(round).join(",")]);
  }
  if (spec.lineStart) {
    common.push(["head", XFDF_LINE_ENDING[spec.lineStart]]);
  }
  if (spec.lineEnd) {
    common.push(["tail", XFDF_LINE_ENDING[spec.lineEnd]]);
  }
  if (spec.vertices && spec.vertices.length >= 4) {
    const pairs: string[] = [];
    for (let index = 0; index + 1 < spec.vertices.length; index += 2) {
      pairs.push(
        `${round(spec.vertices[index])},${round(spec.vertices[index + 1])}`,
      );
    }
    common.push(["vertices", pairs.join(";")]);
  }
  if (spec.interiorColor) {
    common.push(["interior-color", hexColor(spec.interiorColor)]);
  }
  if (spec.borderWidth !== undefined) {
    common.push(["width", round(spec.borderWidth)]);
  }
  if (spec.opacity !== undefined && spec.opacity < 1) {
    common.push(["opacity", round(spec.opacity)]);
  }
  if (spec.shapeFillOpacity !== undefined && spec.shapeFillOpacity !== 1) {
    common.push(["mrai:fill-opacity", round(spec.shapeFillOpacity)]);
  }
  if (spec.shapeCornerRadius !== undefined && spec.shapeCornerRadius !== 0) {
    common.push(["mrai:corner-radius", round(spec.shapeCornerRadius)]);
  }
  const out = [`    <${element} ${attrs(common)}>`];
  if (spec.contents) {
    out.push(`      <contents>${escapeXml(spec.contents)}</contents>`);
  }
  if (spec.freeText?.defaultAppearance) {
    out.push(
      `      <defaultappearance>${escapeXml(spec.freeText.defaultAppearance)}</defaultappearance>`,
    );
  }
  if (spec.freeText?.defaultStyle) {
    out.push(
      `      <defaultstyle>${escapeXml(spec.freeText.defaultStyle)}</defaultstyle>`,
    );
  }
  if (spec.inkList && spec.inkList.length > 0) {
    // Acrobat's own shape: one <gesture> per stroke, points as "x,y;x,y".
    out.push("      <inklist>");
    for (const stroke of spec.inkList) {
      const pairs: string[] = [];
      for (let i = 0; i + 1 < stroke.length; i += 2) {
        pairs.push(`${round(stroke[i])},${round(stroke[i + 1])}`);
      }
      out.push(`        <gesture>${pairs.join(";")}</gesture>`);
    }
    out.push("      </inklist>");
  }
  out.push(`    </${element}>`);
  for (const reply of spec.replies ?? []) {
    if (!reply.contents) continue;
    const replyAttrs = attrs([
      ["page", String(spec.pageIndex)],
      ["rect", spec.rect.map((n) => round(n)).join(",")],
      ["name", reply.name],
      ["title", reply.author],
      ["inreplyto", spec.name],
      ["replyType", "R"],
      ["date", reply.modifiedDate ?? options.date],
      ["creationdate", reply.creationDate],
    ]);
    out.push(`    <text ${replyAttrs}>`);
    out.push(`      <contents>${escapeXml(reply.contents)}</contents>`);
    out.push("    </text>");
  }
  return out;
}

function round(n: number): string {
  return String(Math.round((Number.isFinite(n) ? n : 0) * 1000) / 1000);
}

/** One annotation read out of an XFDF file. */
export interface ParsedXfdfAnnotation {
  subtype: EditableAnnotationSubtype;
  pageIndex: number;
  rect: [number, number, number, number];
  color?: [number, number, number];
  contents?: string;
  author?: string;
  name?: string;
  quadPoints?: number[];
  /** `name` of the annotation this one replies to. */
  inReplyTo?: string;
  /** One flat [x y x y …] array per ink stroke. */
  inkList?: number[][];
  /** The two endpoints of a Line annotation. */
  line?: [number, number, number, number];
  /** Alternating user-space vertices for Polygon / PolyLine. */
  vertices?: number[];
  /** XFDF line-ending names; retained so unsupported arrows fail closed. */
  lineStartStyle?: string;
  lineEndStyle?: string;
  /** Border-effect attributes used by cloudy Square/Polygon annotations. */
  borderEffectStyle?: string;
  borderEffectIntensity?: number;
  borderDashes?: number[];
  interiorColor?: [number, number, number];
  borderWidth?: number;
  opacity?: number;
  shapeFillOpacity?: number;
  shapeCornerRadius?: number;
}

export interface ParsedXfdf {
  valid: boolean;
  rejectedReason?:
    | "unsupported"
    | "limit"
    | "unsafe_xml"
    | "malformed"
    | "duplicate_identity";
  /** Top-level annotations, replies already attached to their parent. */
  annotations: Array<ParsedXfdfAnnotation & { replies: EditableAnnotationReply[] }>;
  /** Replies whose parent is not in the same file — kept, never dropped. */
  orphanReplies: ParsedXfdfAnnotation[];
  /** Annotation elements recognized as XFDF but not supported by this reader. */
  unsupportedAnnotationCount: number;
  href?: string;
  sourceSha256?: string;
}

export type XfdfDocumentMatch = "match" | "mismatch" | "unbound";

function filenameFromHref(value: string): string | null {
  if (
    !value ||
    value.length > XFDF_MAX_HREF_CODE_UNITS ||
    XFDF_SOURCE_ID_UNSAFE.test(value)
  ) {
    return null;
  }
  const path = value.split(/[?#]/, 1)[0].replace(/\\/g, "/");
  const encoded = path.slice(path.lastIndexOf("/") + 1);
  if (!encoded) return null;
  try {
    return decodeURIComponent(encoded).normalize("NFC");
  } catch {
    return null;
  }
}

/** Compare every identity XFDF can honestly carry, strongest proof first. */
export function matchXfdfDocument(
  parsed: Pick<ParsedXfdf, "href" | "sourceSha256">,
  current: { filename?: string; sourceSha256?: string | null },
): XfdfDocumentMatch {
  const incomingSha256 = parsed.sourceSha256?.trim().toLowerCase();
  const currentSha256 = current.sourceSha256?.trim().toLowerCase();
  if (incomingSha256 && SHA256_PATTERN.test(incomingSha256) && currentSha256) {
    return SHA256_PATTERN.test(currentSha256) &&
      incomingSha256 === currentSha256
      ? "match"
      : "mismatch";
  }
  if (parsed.href && current.filename) {
    const incomingName = filenameFromHref(parsed.href);
    const currentName = filenameFromHref(current.filename);
    if (!incomingName || !currentName) return "mismatch";
    return incomingName === currentName ? "match" : "mismatch";
  }
  return "unbound";
}

/**
 * Parse an XFDF document. Unknown elements are ignored rather than rejected: a
 * file from Acrobat routinely carries stamps, ink and free text we do not model
 * yet, and refusing the whole file over one unsupported comment would be worse
 * than importing the rest.
 */
export function fromXfdf(xml: string): ParsedXfdf {
  const rejected = (
    rejectedReason: NonNullable<ParsedXfdf["rejectedReason"]>,
  ): ParsedXfdf => ({
    valid: false,
    rejectedReason,
    annotations: [],
    orphanReplies: [],
    unsupportedAnnotationCount: 0,
  });
  if (typeof DOMParser === "undefined") return rejected("unsupported");
  if (xml.length > XFDF_MAX_CODE_UNITS) return rejected("limit");
  if (/<!DOCTYPE|<!ENTITY/i.test(xml)) return rejected("unsafe_xml");
  let doc: Document;
  try {
    doc = new DOMParser().parseFromString(xml, "application/xml");
  } catch {
    return rejected("malformed");
  }
  if (doc.getElementsByTagName("parsererror").length > 0) {
    return rejected("malformed");
  }
  const root = doc.documentElement;
  if (
    !root ||
    root.localName !== "xfdf" ||
    (root.namespaceURI !== null && root.namespaceURI !== XFDF_NAMESPACE)
  ) {
    return rejected("malformed");
  }

  const rootChildren = Array.from(root.children);
  const fileNodes = rootChildren.filter((node) => node.localName === "f");
  const annotationContainers = rootChildren.filter(
    (node) => node.localName === "annots",
  );
  const sourceNodes = rootChildren.filter(
    (node) =>
      node.localName === "source" && node.namespaceURI === MRAI_XFDF_NAMESPACE,
  );
  if (
    fileNodes.length > 1 ||
    annotationContainers.length > 1 ||
    sourceNodes.length > 1
  ) {
    return rejected("malformed");
  }
  const href = fileNodes[0]?.getAttribute("href") ?? undefined;
  if (
    href &&
    (href.length > XFDF_MAX_HREF_CODE_UNITS || XFDF_SOURCE_ID_UNSAFE.test(href))
  ) {
    return rejected("limit");
  }
  const rawSourceSha256 = sourceNodes[0]?.getAttribute("sha256") ?? undefined;
  const sourceSha256 = rawSourceSha256?.trim().toLowerCase();
  if (
    sourceNodes.length === 1 &&
    (!sourceSha256 || !SHA256_PATTERN.test(sourceSha256))
  ) {
    return rejected("malformed");
  }
  const parsed: ParsedXfdfAnnotation[] = [];
  let unsupportedAnnotationCount = 0;
  const seenNames = new Set<string>();
  const container = annotationContainers[0];
  const children = container ? Array.from(container.children) : [];
  if (children.length > XFDF_MAX_ANNOTATIONS) return rejected("limit");
  for (const node of children) {
    const subtype = SUBTYPE_FOR[node.localName.toLowerCase()];
    if (!subtype) {
      unsupportedAnnotationCount += 1;
      continue;
    }
    const rawName = node.getAttribute("name");
    const name = rawName === null ? undefined : normalizeXfdfSourceId(rawName);
    const rawInReplyTo = node.getAttribute("inreplyto");
    const inReplyTo =
      rawInReplyTo === null
        ? undefined
        : normalizeXfdfSourceId(rawInReplyTo);
    if (
      (rawName !== null && !name) ||
      (rawInReplyTo !== null && !inReplyTo)
    ) {
      return rejected("malformed");
    }
    if (name) {
      if (seenNames.has(name)) return rejected("duplicate_identity");
      seenNames.add(name);
    }
    const rawRect = node.getAttribute("rect");
    const rect = numberList(rawRect, 4);
    const page = Number(node.getAttribute("page"));
    if (
      rawRect === null ||
      rect.length !== 4 ||
      !Number.isSafeInteger(page) ||
      page < 0
    ) {
      return rejected("malformed");
    }
    const contents =
      node.getElementsByTagName("contents")[0]?.textContent ?? undefined;
    const author = node.getAttribute("title") ?? undefined;
    if (
      (contents?.length ?? 0) > XFDF_MAX_CONTENT_CODE_UNITS ||
      (author?.length ?? 0) > XFDF_MAX_AUTHOR_CODE_UNITS
    ) {
      return rejected("limit");
    }
    const rawQuad = node.getAttribute("coords");
    const rawStart = node.getAttribute("start");
    const rawEnd = node.getAttribute("end");
    const rawVertices = node.getAttribute("vertices");
    const quad = numberList(rawQuad, 32_000);
    const start = numberList(rawStart, 2);
    const end = numberList(rawEnd, 2);
    // Managed Polygon geometry is bounded to 200 points. Reject a larger
    // attribute before converting it into a second unbounded numeric array.
    const vertices = numberList(rawVertices, 400);
    if (
      (rawQuad !== null && (quad.length < 8 || quad.length % 8 !== 0)) ||
      (rawStart !== null && start.length !== 2) ||
      (rawEnd !== null && end.length !== 2) ||
      (rawVertices !== null &&
        (vertices.length < 4 || vertices.length % 2 !== 0))
    ) {
      return rejected("malformed");
    }
    const gestureNodes = Array.from(node.getElementsByTagName("gesture"));
    if (gestureNodes.length > XFDF_MAX_INK_GESTURES) return rejected("limit");
    const gestures = gestureNodes.map((gesture) =>
      numberList(
        gesture.textContent ?? "",
        XFDF_MAX_INK_VALUES_PER_GESTURE,
      ),
    );
    if (
      gestures.some(
        (stroke) => stroke.length < 4 || stroke.length % 2 !== 0,
      )
    ) {
      return rejected("malformed");
    }
    const rawWidth = node.getAttribute("width");
    const rawOpacity = node.getAttribute("opacity");
    const hasShapeFillOpacity = node.hasAttributeNS(
      MRAI_XFDF_NAMESPACE,
      "fill-opacity",
    );
    const hasShapeCornerRadius = node.hasAttributeNS(
      MRAI_XFDF_NAMESPACE,
      "corner-radius",
    );
    const rawShapeFillOpacity = hasShapeFillOpacity
      ? node.getAttributeNS(MRAI_XFDF_NAMESPACE, "fill-opacity")
      : null;
    const rawShapeCornerRadius = hasShapeCornerRadius
      ? node.getAttributeNS(MRAI_XFDF_NAMESPACE, "corner-radius")
      : null;
    const rawLineStartStyle = node.getAttribute("head");
    const rawLineEndStyle = node.getAttribute("tail");
    const rawBorderEffectStyle = node.getAttribute("style");
    const rawBorderEffectIntensity = node.getAttribute("intensity");
    const rawBorderDashes = node.getAttribute("dashes");
    const widthAttr = Number(rawWidth);
    const opacityAttr = Number(rawOpacity);
    const shapeFillOpacity = Number(rawShapeFillOpacity);
    const shapeCornerRadius = Number(rawShapeCornerRadius);
    const borderEffectIntensity = Number(rawBorderEffectIntensity);
    const borderDashes = numberList(rawBorderDashes, 32);
    const rawColor = node.getAttribute("color");
    const rawInteriorColor = node.getAttribute("interior-color");
    const color = parseHexColor(rawColor);
    const interiorColor = parseHexColor(rawInteriorColor);
    if (
      (rawWidth !== null &&
        (!rawWidth.trim() || !Number.isFinite(widthAttr) || widthAttr < 0)) ||
      (rawOpacity !== null &&
        (!rawOpacity.trim() ||
          !Number.isFinite(opacityAttr) ||
          opacityAttr < 0 ||
          opacityAttr > 1)) ||
      (hasShapeFillOpacity &&
        (rawShapeFillOpacity === null ||
          !rawShapeFillOpacity.trim() ||
          !Number.isFinite(shapeFillOpacity) ||
          shapeFillOpacity < 0 ||
          shapeFillOpacity > 1 ||
          interiorColor === null ||
          (subtype !== "Square" && subtype !== "Circle"))) ||
      (hasShapeCornerRadius &&
        (rawShapeCornerRadius === null ||
          !rawShapeCornerRadius.trim() ||
          !Number.isFinite(shapeCornerRadius) ||
          shapeCornerRadius <= 0 ||
          shapeCornerRadius > 50 ||
          subtype !== "Square")) ||
      (rawColor !== null && !color) ||
      (rawInteriorColor !== null && !interiorColor) ||
      [rawLineStartStyle, rawLineEndStyle, rawBorderEffectStyle].some(
        (value) =>
          value !== null &&
          (!value.trim() ||
            value.length > 64 ||
            XFDF_SOURCE_ID_UNSAFE.test(value)),
      ) ||
      (rawBorderEffectIntensity !== null &&
        (!rawBorderEffectIntensity.trim() ||
          !Number.isFinite(borderEffectIntensity) ||
          borderEffectIntensity < 0)) ||
      (rawBorderDashes !== null && borderDashes.length === 0)
    ) {
      return rejected("malformed");
    }
    parsed.push({
      subtype,
      pageIndex: page,
      rect: [rect[0], rect[1], rect[2], rect[3]],
      color: color ?? undefined,
      contents: contents?.trim() ? contents : undefined,
      author,
      name,
      quadPoints: quad.length >= 8 ? quad : undefined,
      inReplyTo,
      inkList: gestures.length > 0 ? gestures : undefined,
      line:
        start.length === 2 && end.length === 2
          ? [start[0], start[1], end[0], end[1]]
          : undefined,
      vertices:
        vertices.length >= 4 && vertices.length % 2 === 0
          ? vertices
          : undefined,
      lineStartStyle: rawLineStartStyle?.trim() || undefined,
      lineEndStyle: rawLineEndStyle?.trim() || undefined,
      borderEffectStyle: rawBorderEffectStyle?.trim() || undefined,
      borderEffectIntensity:
        rawBorderEffectIntensity !== null ? borderEffectIntensity : undefined,
      borderDashes: rawBorderDashes !== null ? borderDashes : undefined,
      interiorColor: interiorColor ?? undefined,
      borderWidth: rawWidth !== null ? widthAttr : undefined,
      opacity: rawOpacity !== null ? opacityAttr : undefined,
      shapeFillOpacity: hasShapeFillOpacity ? shapeFillOpacity : undefined,
      shapeCornerRadius: hasShapeCornerRadius ? shapeCornerRadius : undefined,
    });
  }

  const byName = new Map<string, ParsedXfdfAnnotation & { replies: EditableAnnotationReply[] }>();
  const annotations: Array<ParsedXfdfAnnotation & { replies: EditableAnnotationReply[] }> = [];
  for (const item of parsed) {
    if (item.inReplyTo) continue;
    const withReplies = { ...item, replies: [] as EditableAnnotationReply[] };
    annotations.push(withReplies);
    if (item.name) byName.set(item.name, withReplies);
  }
  const orphanReplies: ParsedXfdfAnnotation[] = [];
  for (const item of parsed) {
    if (!item.inReplyTo) continue;
    const parent = byName.get(item.inReplyTo);
    if (!parent) {
      orphanReplies.push(item);
      continue;
    }
    parent.replies.push({
      name: item.name,
      author: item.author,
      contents: item.contents ?? "",
    });
  }
  return {
    valid: true,
    annotations,
    orphanReplies,
    unsupportedAnnotationCount,
    href,
    ...(sourceSha256 ? { sourceSha256 } : null),
  };
}

/**
 * Merge parsed annotations onto existing specs, keyed by name. An incoming
 * annotation with a name we already hold REPLACES it; one we have never seen is
 * appended. Importing the same file twice therefore lands the same document,
 * which is the whole point of carrying /NM through the round trip.
 */
export function mergeXfdfIntoSpecs(
  existing: ReadonlyArray<EditableAnnotationSpec>,
  incoming: ParsedXfdf,
  rebuild: (
    parsed: ParsedXfdfAnnotation & { replies: EditableAnnotationReply[] },
    previous?: EditableAnnotationSpec,
  ) => EditableAnnotationSpec | null,
): EditableAnnotationSpec[] {
  const out = [...existing];
  const indexByName = new Map<string, number>();
  out.forEach((spec, index) => {
    if (spec.name) indexByName.set(spec.name, index);
  });
  for (const parsed of incoming.annotations) {
    const at = parsed.name ? indexByName.get(parsed.name) : undefined;
    const built = rebuild(parsed, at === undefined ? undefined : out[at]);
    if (!built) continue;
    if (at === undefined) {
      if (built.name) indexByName.set(built.name, out.length);
      out.push(built);
    } else {
      out[at] = built;
    }
  }
  return out;
}
