/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { JsonValue } from "@mrnote/core/web/lib/collab/json-crdt.ts";
import { readPdfCommandOwnershipManifest } from "../command-ownership";
import type { PdfEngineCommand, PdfEngineSourceIdentity } from "../contracts";
import type { PdfNativeContentAppendCategory, PdfNativeContentAppendKind, PdfNativeContentAppendReusedResource } from "./content-stream-transaction";
import { MRAI_NATIVE_CONTENT_APPEND_ACTION, MRAI_NATIVE_CONTENT_APPEND_SCHEMA_VERSION } from "./vocabulary";
import type { PdfNativeContentAppendColor, PdfNativeContentAppendContent, PdfNativeContentAppendMatrix, PdfNativeContentAppendPageGeometry, PdfNativeContentAppendPathCommand, PdfNativeContentAppendPayload, PdfNativeContentAppendRequest, PdfNativeContentAppendSpan, PdfNativeShapedTextPiece } from "./content-append-transaction";
import { PDF_COMPLEX_SCRIPT_FONTS } from "../../unicode-font-asset";

export const PDF_NATIVE_CONTENT_APPEND_CONTROL_IDS = Object.freeze({
  text: "PDFCTL.PAGE.CONTENT.APPEND_TEXT",
  image: "PDFCTL.PAGE.CONTENT.APPEND_IMAGE",
  path: "PDFCTL.PAGE.CONTENT.APPEND_PATH",
} as const);

export const PDF_NATIVE_CONTENT_APPEND_FEATURE_IDS = Object.freeze({
  text: "PDF-INS-CONTENT-01",
  image: "PDF-INS-CONTENT-02",
  path: "PDF-CMT-DRAW-02",
} as const);

export const SHA256_PATTERN = /^[0-9a-f]{64}$/;

export const RESOURCE_NAME_PATTERN = /^[A-Za-z][A-Za-z0-9_.+\-]{0,63}$/;

export const MAX_TEXT_CODE_UNITS = 2_000;

export const MAX_PATH_COMMANDS = 4_096;

export const MAX_ABS_COORDINATE = 10_000_000;

export const CONTROL_CHARACTERS = /[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/u;

/** Text-space rotation (a b c d) that keeps text upright on a /Rotate page. */
export const TEXT_ROTATION_MATRICES: Readonly<
  Record<0 | 90 | 180 | 270, readonly [number, number, number, number]>
> = Object.freeze({
  0: [1, 0, 0, 1],
  90: [0, 1, -1, 0],
  180: [-1, 0, 0, -1],
  270: [0, -1, 1, 0],
});

export class PdfNativeContentAppendError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "PdfNativeContentAppendError";
  }
}

export function invalid(message: string): never {
  throw new PdfNativeContentAppendError(message);
}

export function plainRecord(value: unknown): Readonly<Record<string, unknown>> | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) return null;
  return value as Readonly<Record<string, unknown>>;
}

export function finite(value: unknown, label: string): number {
  if (
    typeof value !== "number" ||
    !Number.isFinite(value) ||
    Math.abs(value) > MAX_ABS_COORDINATE
  ) {
    invalid(`${label} is not a finite PDF number.`);
  }
  return value;
}

export function unit(value: unknown, label: string): number {
  const number = finite(value, label);
  if (number < 0 || number > 1) invalid(`${label} must lie in [0, 1].`);
  return number;
}

export function exactSha(value: unknown, label: string): string {
  if (typeof value !== "string" || !SHA256_PATTERN.test(value)) {
    invalid(`${label} SHA-256 is invalid.`);
  }
  return value;
}

export function exactKind(value: unknown): PdfNativeContentAppendKind {
  if (value !== "text" && value !== "image" && value !== "path") {
    invalid("Content append kind is invalid.");
  }
  return value;
}

export function readColor(value: unknown, label: string): PdfNativeContentAppendColor {
  const record = plainRecord(value);
  if (!record) invalid(`${label} color is invalid.`);
  const cmyk = record.cmyk;
  if (cmyk !== undefined && (!Array.isArray(cmyk) || cmyk.length !== 4)) {
    invalid(`${label} CMYK color is invalid.`);
  }
  return Object.freeze({
    r: unit(record.r, `${label} red`),
    g: unit(record.g, `${label} green`),
    b: unit(record.b, `${label} blue`),
    ...(cmyk !== undefined
      ? {
          cmyk: Object.freeze(
            ["cyan", "magenta", "yellow", "black"].map((name, index) =>
              unit(cmyk[index], `${label} ${name}`),
            ) as [number, number, number, number],
          ),
        }
      : {}),
  });
}

export function readMatrix(value: unknown, label: string): PdfNativeContentAppendMatrix {
  if (!Array.isArray(value) || value.length !== 6) {
    invalid(`${label} matrix is invalid.`);
  }
  const values = value.map((item, index) => finite(item, `${label}[${index}]`));
  const determinant = values[0] * values[3] - values[1] * values[2];
  if (!Number.isFinite(determinant) || Math.abs(determinant) < 1e-9) {
    invalid(`${label} matrix is singular.`);
  }
  return Object.freeze([
    values[0],
    values[1],
    values[2],
    values[3],
    values[4],
    values[5],
  ] as const);
}

export function readReuse(
  value: unknown,
): readonly PdfNativeContentAppendReusedResource[] {
  if (!Array.isArray(value) || value.length > 8) {
    invalid("Content append reuse list is invalid.");
  }
  return Object.freeze(
    value.map((item) => {
      const record = plainRecord(item);
      const reference = record ? plainRecord(record.reference) : null;
      if (
        !record ||
        !reference ||
        (record.category !== "Font" &&
          record.category !== "XObject" &&
          record.category !== "ExtGState") ||
        typeof record.resourceName !== "string" ||
        !RESOURCE_NAME_PATTERN.test(record.resourceName) ||
        !Number.isSafeInteger(reference.objectNumber) ||
        (reference.objectNumber as number) < 1 ||
        !Number.isSafeInteger(reference.generation) ||
        (reference.generation as number) < 0
      ) {
        invalid("Content append reused resource is invalid.");
      }
      return Object.freeze({
        category: record.category as PdfNativeContentAppendCategory,
        resourceName: record.resourceName,
        reference: Object.freeze({
          objectNumber: reference.objectNumber as number,
          generation: reference.generation as number,
        }),
        expectedObjectSha256: exactSha(
          record.expectedObjectSha256,
          "Content append reused resource",
        ),
      });
    }),
  );
}

export function readSpan(value: unknown): PdfNativeContentAppendSpan {
  const record = plainRecord(value);
  if (
    !record ||
    !Number.isSafeInteger(record.start) ||
    !Number.isSafeInteger(record.end) ||
    (record.start as number) < 0 ||
    (record.end as number) <= (record.start as number)
  ) {
    invalid("Content append span is invalid.");
  }
  return Object.freeze({
    start: record.start as number,
    end: record.end as number,
    expectedSha256: exactSha(record.expectedSha256, "Content append span"),
  });
}

export function readPathCommands(
  value: unknown,
): readonly PdfNativeContentAppendPathCommand[] {
  if (!Array.isArray(value) || value.length < 1 || value.length > MAX_PATH_COMMANDS) {
    invalid("Content append path commands are invalid.");
  }
  return Object.freeze(
    value.map((item): PdfNativeContentAppendPathCommand => {
      const record = plainRecord(item);
      if (!record) invalid("Content append path command is invalid.");
      switch (record.op) {
        case "m":
        case "l":
          return Object.freeze({
            op: record.op,
            x: finite(record.x, "path x"),
            y: finite(record.y, "path y"),
          });
        case "c":
          return Object.freeze({
            op: "c" as const,
            x1: finite(record.x1, "path x1"),
            y1: finite(record.y1, "path y1"),
            x2: finite(record.x2, "path x2"),
            y2: finite(record.y2, "path y2"),
            x: finite(record.x, "path x"),
            y: finite(record.y, "path y"),
          });
        case "re":
          return Object.freeze({
            op: "re" as const,
            x: finite(record.x, "path x"),
            y: finite(record.y, "path y"),
            width: finite(record.width, "path width"),
            height: finite(record.height, "path height"),
          });
        case "h":
          return Object.freeze({ op: "h" as const });
        default:
          return invalid("Content append path operator is invalid.");
      }
    }),
  );
}

/** The API's shaping of each line (PdfNativeShapedTextPiece): every line's
 * pieces hold its characters exactly, a shaped piece names a packaged face
 * by its pinned digest and has glyphs, each glyph's cluster inside its text. */
function readShapedLines(value: unknown, text: string): readonly (readonly PdfNativeShapedTextPiece[])[] {
  const lines = text.split(/\r?\n/);
  if (!Array.isArray(value) || value.length !== lines.length) invalid("Content append shaped lines are invalid.");
  const sorted = (characters: string) => [...characters].sort().join("");
  return Object.freeze(value.map((line, index) => {
    if (!Array.isArray(line) || line.length > 400) invalid("Content append shaped line is invalid.");
    const pieces = line.map((entry) => {
      const piece = plainRecord(entry);
      if (!piece || typeof piece.text !== "string" || !piece.text || typeof piece.rtl !== "boolean" ||
        (piece.kind !== "plain" && piece.kind !== "shaped")) invalid("Content append shaped piece is invalid.");
      if (piece.kind === "plain") return Object.freeze({ text: piece.text, rtl: piece.rtl, kind: "plain" as const });
      const face = typeof piece.sourceId === "string" ? PDF_COMPLEX_SCRIPT_FONTS[piece.sourceId] : undefined;
      const glyphs = piece.glyphs;
      if (!face || piece.fontSha256 !== face.sha256 || !Number.isSafeInteger(piece.unitsPerEm) ||
        (piece.unitsPerEm as number) <= 0 || !Array.isArray(glyphs) || glyphs.length === 0 ||
        glyphs.length > MAX_TEXT_CODE_UNITS * 4 ||
        glyphs.some((glyph) => !Array.isArray(glyph) || glyph.length !== 6 ||
          glyph.some((part) => !Number.isSafeInteger(part) || Math.abs(part as number) > 1_000_000) ||
          (glyph[0] as number) <= 0 || (glyph[1] as number) < 0 || (glyph[1] as number) >= (piece.text as string).length)) {
        invalid("Content append shaped glyphs are invalid.");
      }
      return Object.freeze({
        text: piece.text, rtl: piece.rtl, kind: "shaped" as const, sourceId: piece.sourceId as string,
        fontSha256: face.sha256, unitsPerEm: piece.unitsPerEm as number,
        glyphs: Object.freeze((glyphs as number[][]).map((glyph) => Object.freeze([...glyph]) as unknown as
          readonly [number, number, number, number, number, number])),
      });
    });
    if (sorted(pieces.map((piece) => piece.text).join("")) !== sorted(lines[index])) {
      invalid("Content append shaped line does not hold its text.");
    }
    return Object.freeze(pieces);
  }));
}

export function readContent(value: unknown): PdfNativeContentAppendContent {
  const record = plainRecord(value);
  if (!record) invalid("Content append content is invalid.");
  if (record.kind === "text") {
    if (
      typeof record.text !== "string" ||
      !record.text.trim() ||
      record.text.length > MAX_TEXT_CODE_UNITS ||
      CONTROL_CHARACTERS.test(record.text)
    ) {
      invalid("Content append text is empty, too long or contains control characters.");
    }
    const fontSizePoints = finite(record.fontSizePoints, "font size");
    const lineHeightPoints = finite(record.lineHeightPoints, "line height");
    if (fontSizePoints <= 0 || lineHeightPoints <= 0) {
      invalid("Content append font size and line height must be positive.");
    }
    const rotation = record.rotation === undefined ? 0 : record.rotation;
    if (rotation !== 0 && rotation !== 90 && rotation !== 180 && rotation !== 270) {
      invalid("Content append text rotation is invalid.");
    }
    const shapedLines = record.shapedLines === undefined ? undefined
      : readShapedLines(record.shapedLines, record.text);
    if (record.bold !== undefined && typeof record.bold !== "boolean") {
      invalid("Content append text weight is invalid.");
    }
    return Object.freeze({
      kind: "text" as const,
      text: record.text,
      fontSizePoints,
      color: readColor(record.color, "text"),
      x: finite(record.x, "text x"),
      top: finite(record.top, "text top"),
      rotation,
      lineHeightPoints,
      ...(record.bold === true ? { bold: true as const } : {}),
      ...(shapedLines ? { shapedLines } : {}),
    });
  }
  if (record.kind === "image") {
    if (
      (record.mimeType !== "image/png" && record.mimeType !== "image/jpeg") ||
      typeof record.imageBase64 !== "string" ||
      !record.imageBase64
    ) {
      invalid("Content append image payload is invalid.");
    }
    return Object.freeze({
      kind: "image" as const,
      mimeType: record.mimeType,
      imageBase64: record.imageBase64,
      matrix: readMatrix(record.matrix, "image"),
    });
  }
  if (record.kind === "path") {
    const strokeRecord = record.stroke === null ? null : plainRecord(record.stroke);
    const fillRecord = record.fill === null ? null : plainRecord(record.fill);
    if (strokeRecord === undefined || fillRecord === undefined) {
      invalid("Content append path style is invalid.");
    }
    if (record.stroke !== null && !strokeRecord) {
      invalid("Content append stroke is invalid.");
    }
    if (record.fill !== null && !fillRecord) {
      invalid("Content append fill is invalid.");
    }
    const stroke = strokeRecord
      ? (() => {
          const width = finite(strokeRecord.width, "stroke width");
          if (width < 0) invalid("Content append stroke width is negative.");
          const dash =
            strokeRecord.dash === null || strokeRecord.dash === undefined
              ? null
              : Array.isArray(strokeRecord.dash) &&
                  strokeRecord.dash.length >= 1 &&
                  strokeRecord.dash.length <= 8
                ? Object.freeze(
                    strokeRecord.dash.map((item) => {
                      const number = finite(item, "dash");
                      if (number < 0) invalid("Content append dash is negative.");
                      return number;
                    }),
                  )
                : invalid("Content append dash array is invalid.");
          if (dash && dash.every((item) => item === 0)) {
            invalid("Content append dash array is all zero.");
          }
          const cap = strokeRecord.cap ?? 0;
          const join = strokeRecord.join ?? 0;
          if (![0, 1, 2].includes(cap as number) || ![0, 1, 2].includes(join as number)) {
            invalid("Content append line cap or join is invalid.");
          }
          return Object.freeze({
            color: readColor(strokeRecord.color, "stroke"),
            width,
            dash,
            cap: cap as 0 | 1 | 2,
            join: join as 0 | 1 | 2,
          });
        })()
      : null;
    const fill = fillRecord
      ? Object.freeze({ color: readColor(fillRecord.color, "fill") })
      : null;
    const opacity =
      record.opacity === null || record.opacity === undefined
        ? null
        : unit(record.opacity, "opacity");
    const fillOpacity =
      record.fillOpacity === null || record.fillOpacity === undefined
        ? null
        : unit(record.fillOpacity, "fill opacity");
    if (!stroke && !fill) invalid("Content append path paints nothing.");
    return Object.freeze({
      kind: "path" as const,
      stroke,
      fill,
      opacity,
      fillOpacity,
      commands: readPathCommands(record.commands),
    });
  }
  return invalid("Content append content kind is invalid.");
}

export function readPdfNativeContentAppendPayload(
  value: unknown,
): PdfNativeContentAppendPayload {
  const record = plainRecord(value);
  const request = record ? plainRecord(record.request) : null;
  if (
    !record ||
    !request ||
    record.schemaVersion !== MRAI_NATIVE_CONTENT_APPEND_SCHEMA_VERSION ||
    !Number.isSafeInteger(record.pageNumber) ||
    (record.pageNumber as number) < 1 ||
    (record.pageNumber as number) > 100_000
  ) {
    invalid("Content append payload is invalid.");
  }
  const sourceSha256 = exactSha(record.sourceSha256, "Content append source");
  const pageNumber = record.pageNumber as number;
  let parsedRequest: PdfNativeContentAppendRequest;
  if (request.direction === "remove") {
    parsedRequest = Object.freeze({
      direction: "remove" as const,
      appendKind: exactKind(request.appendKind),
      span: readSpan(request.span),
    });
  } else if (request.direction === "append" && "blockBase64" in request) {
    if (typeof request.blockBase64 !== "string" || !request.blockBase64) {
      invalid("Content append block is invalid.");
    }
    parsedRequest = Object.freeze({
      direction: "append" as const,
      appendKind: exactKind(request.appendKind),
      blockBase64: request.blockBase64,
      reuse: readReuse(request.reuse),
    });
  } else if (request.direction === "append") {
    parsedRequest = Object.freeze({
      direction: "append" as const,
      content: readContent(request.content),
    });
  } else {
    return invalid("Content append direction is invalid.");
  }
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_CONTENT_APPEND_SCHEMA_VERSION,
    sourceSha256,
    pageNumber,
    request: parsedRequest,
  });
}

/** Displayed page size in PDF units (crop box turned by /Rotate). */
export function pdfNativeContentAppendDisplaySize(
  geometry: PdfNativeContentAppendPageGeometry,
): Readonly<{ width: number; height: number }> {
  const width = geometry.cropBox.xMax - geometry.cropBox.xMin;
  const height = geometry.cropBox.yMax - geometry.cropBox.yMin;
  return geometry.rotation === 90 || geometry.rotation === 270
    ? Object.freeze({ width: height, height: width })
    : Object.freeze({ width, height });
}

/**
 * Map a point given as fractions of the displayed page (origin top-left, as
 * the reader sees it) to PDF user space, honouring the page's /Rotate.
 */
export function pdfNativeContentAppendDisplayPoint(
  geometry: PdfNativeContentAppendPageGeometry,
  fractionX: number,
  fractionY: number,
): Readonly<{ x: number; y: number }> {
  const crop = geometry.cropBox;
  const width = crop.xMax - crop.xMin;
  const height = crop.yMax - crop.yMin;
  switch (geometry.rotation) {
    case 90:
      return Object.freeze({
        x: crop.xMin + fractionY * width,
        y: crop.yMin + fractionX * height,
      });
    case 180:
      return Object.freeze({
        x: crop.xMax - fractionX * width,
        y: crop.yMin + fractionY * height,
      });
    case 270:
      return Object.freeze({
        x: crop.xMax - fractionY * width,
        y: crop.yMax - fractionX * height,
      });
    default:
      return Object.freeze({
        x: crop.xMin + fractionX * width,
        y: crop.yMax - fractionY * height,
      });
  }
}

/**
 * Matrix that places the image unit square over a displayed rectangle
 * (fractions of the displayed page, origin top-left), upright on screen.
 */
export function pdfNativeContentAppendDisplayRectMatrix(
  geometry: PdfNativeContentAppendPageGeometry,
  rect: Readonly<{ x: number; y: number; width: number; height: number }>,
): PdfNativeContentAppendMatrix {
  const size = pdfNativeContentAppendDisplaySize(geometry);
  const w = rect.width * size.width;
  const h = rect.height * size.height;
  // The unit square's origin is the image's bottom-left as seen on screen.
  const origin = pdfNativeContentAppendDisplayPoint(
    geometry,
    rect.x,
    rect.y + rect.height,
  );
  const [a, b, c, d] = TEXT_ROTATION_MATRICES[geometry.rotation];
  return Object.freeze([a * w, b * w, c * h, d * h, origin.x, origin.y] as const);
}

/** Parse `#rgb` / `#rrggbb` into unit RGB; null for anything else. */
export function pdfNativeContentAppendColorFromHex(
  value: string | null | undefined,
): PdfNativeContentAppendColor | null {
  if (typeof value !== "string") return null;
  const match = /^#?([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(value.trim());
  if (!match) return null;
  const hex = match[1].length === 3
    ? match[1].split("").map((digit) => digit + digit).join("")
    : match[1];
  return Object.freeze({
    r: parseInt(hex.slice(0, 2), 16) / 255,
    g: parseInt(hex.slice(2, 4), 16) / 255,
    b: parseInt(hex.slice(4, 6), 16) / 255,
  });
}

export function pdfNativeContentAppendFeatureId(
  kind: PdfNativeContentAppendKind,
): (typeof PDF_NATIVE_CONTENT_APPEND_FEATURE_IDS)[PdfNativeContentAppendKind] {
  return PDF_NATIVE_CONTENT_APPEND_FEATURE_IDS[kind];
}

export function createPdfNativeContentAppendOwnershipManifest(
  kind: PdfNativeContentAppendKind,
) {
  return readPdfCommandOwnershipManifest([
    {
      featureId: pdfNativeContentAppendFeatureId(kind),
      action: MRAI_NATIVE_CONTENT_APPEND_ACTION,
      authority: "mrai-native",
      mutation: "native-bytes",
      output: "current-pdf",
    },
  ]);
}

export function pdfNativeContentAppendPayloadKind(
  payload: PdfNativeContentAppendPayload,
): PdfNativeContentAppendKind {
  const request = payload.request;
  return "content" in request ? request.content.kind : request.appendKind;
}

export function buildPdfNativeContentAppendCommand(input: {
  readonly source: PdfEngineSourceIdentity;
  readonly payload: PdfNativeContentAppendPayload;
  readonly transactionId: string;
  readonly sessionRevision: number;
  readonly selectionRevision: number;
}): PdfEngineCommand {
  const payload = readPdfNativeContentAppendPayload(input.payload);
  if (
    payload.sourceSha256 !== input.source.sourceSha256 ||
    !input.transactionId.trim() ||
    !Number.isSafeInteger(input.sessionRevision) ||
    input.sessionRevision < 0 ||
    !Number.isSafeInteger(input.selectionRevision) ||
    input.selectionRevision < 1
  ) {
    invalid("Content append command authority is stale.");
  }
  return Object.freeze({
    ...input.source,
    featureId: pdfNativeContentAppendFeatureId(pdfNativeContentAppendPayloadKind(payload)),
    controlId: PDF_NATIVE_CONTENT_APPEND_CONTROL_IDS[pdfNativeContentAppendPayloadKind(payload)],
    action: MRAI_NATIVE_CONTENT_APPEND_ACTION,
    transactionId: input.transactionId,
    sessionRevision: input.sessionRevision,
    selectionRevision: input.selectionRevision,
    output: "current-pdf",
    payload: payload as unknown as JsonValue,
  });
}
