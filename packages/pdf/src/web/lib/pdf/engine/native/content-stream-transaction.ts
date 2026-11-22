/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { completePdfLengthBeforeTrailingBytes } from "./complete-pdf-before-trailing-bytes";
import { LATIN1_EXACT } from "./latin1";
import { pdfNativeTextObjectEdgesAt, type PdfNativeTextRuleSpan } from "./text-rule-block";
import { pageInterfaceLocale } from "@mrnote/core/web/lib/i18n/content-locale.ts";
import { isPdfManagedBookmarksTitle, pdfManagedBookmarksTitle } from "../../managed-bookmarks-title";
export type { PdfNativeTextRuleSpan } from "./text-rule-block";
import { concatPdfNativeEmbeddedImageTransformMatrices, pdfNativeEmbeddedImagePaintMatrixMatches, pdfNativeEmbeddedImagePlacement, type PdfNativeEmbeddedImageCropMode } from "./embedded-image-model";
export { concatPdfNativeEmbeddedImageTransformMatrices, pdfNativeEmbeddedImagePlacement };
import { EMBEDDED_IMAGE_FRAME_MIDDLE, EMBEDDED_IMAGE_FRAME_WS, EMBEDDED_IMAGE_TRANSFORM_MAX_OPERATOR_BYTES, EMBEDDED_IMAGE_TRANSFORM_NUMBER, PDF_EXISTING_RESOURCE_NAME, PDF_RESOURCE_NAME, decodeLatin1, exactEmbeddedImageTransformClip, exactEmbeddedImageTransformMatrix, fail, parsePdfNativeEmbeddedImageFrameClip, parsePdfNativeEmbeddedImageTransformOperator, readPdfNativeEmbeddedImageTransformEvidence } from "./embedded-image-model";
export { parsePdfNativeEmbeddedImageFrameClip, parsePdfNativeEmbeddedImageTransformOperator, readPdfNativeEmbeddedImageTransformEvidence };
import {
  MRAI_NATIVE_CONTENT_APPEND_SCHEMA_VERSION,
  MRAI_NATIVE_CONTENT_STREAM_SCHEMA_VERSION,
  MRAI_NATIVE_EMBEDDED_IMAGE_REMOVE_SCHEMA_VERSION,
  MRAI_NATIVE_EMBEDDED_IMAGE_REPLACE_SCHEMA_VERSION,
  MRAI_NATIVE_EMBEDDED_IMAGE_TRANSFORM_SCHEMA_VERSION,
  MRAI_NATIVE_OUTLINE_SCHEMA_VERSION,
  MRAI_NATIVE_PAGE_BACKGROUND_SCHEMA_VERSION,
  MRAI_NATIVE_PAGE_SPLIT_REFUSALS,
  MRAI_NATIVE_PAGE_SPLIT_SCHEMA_VERSION,
  PDF_NATIVE_PAGE_BACKGROUND_MAX_PAGES,
  PDF_NATIVE_PAGE_SPLIT_MAX_LINES,
  PDF_NATIVE_PAGE_SPLIT_MAX_PAGES,
  PDF_NATIVE_PAGE_SPLIT_MAX_PARTS,
} from "./vocabulary";
export {
  MRAI_NATIVE_CONTENT_APPEND_SCHEMA_VERSION,
  MRAI_NATIVE_CONTENT_STREAM_SCHEMA_VERSION,
  MRAI_NATIVE_EMBEDDED_IMAGE_REMOVE_SCHEMA_VERSION,
  MRAI_NATIVE_EMBEDDED_IMAGE_REPLACE_SCHEMA_VERSION,
  MRAI_NATIVE_EMBEDDED_IMAGE_TRANSFORM_SCHEMA_VERSION,
  MRAI_NATIVE_OUTLINE_SCHEMA_VERSION,
};
import {
  MAX_MRAI_NATIVE_EMBEDDED_IMAGE_BYTES,
  MAX_MRAI_NATIVE_SOURCE_BYTES,
  PdfNativeTransactionError,
  sha256PdfBytes,
  type PdfNativeTransactionErrorCode,
} from "./transaction-error";
export {
  MAX_MRAI_NATIVE_EMBEDDED_IMAGE_BYTES,
  MAX_MRAI_NATIVE_SOURCE_BYTES,
  PdfNativeTransactionError,
  sha256PdfBytes,
  type PdfNativeTransactionErrorCode,
};
import type { PdfNativeIndirectObjectRef } from "./list-text-fidelity";
import {
  MRAI_NATIVE_BOX_MARK_CLOSE,
  MRAI_NATIVE_BOX_MARK_OPEN,
  mraiNativeBoxMarkIssue,
  type PdfNativeBoxMark,
} from "./box-mark";
import { Inflate, inflateSync, Unzlib, zlibSync } from "fflate";

import {
  PDF_BLEND_MODE_BY_PDF_NAME,
  PDF_NAME_BY_BLEND_MODE,
  isPdfStandardBlendMode,
  type PdfStandardBlendMode,
} from "../../blend-modes";

import {
  MANAGED_OUTLINE_MAX_ABS_COORDINATE,
  MANAGED_OUTLINE_MAX_ZOOM,
  OUTLINE_TITLE_MAX_LENGTH,
  hasUnpairedUtf16Surrogate,
  sanitizeOutlineTitle,
} from "@mrnote/core/web/lib/pdf/outline.ts";
import {
  MRAI_NATIVE_FONT_EXTENSION_PACKAGE_SCHEMA_VERSION,
  buildPdfNativeFontExtensionPackage,
  type PdfNativeFontExtensionPackage,
} from "./font-extension-package";
import { isPdfComplexScriptFontSha256, PDF_UNICODE_FALLBACK_FONT_SHA256 } from "../../unicode-font-asset";
import {
  isPdfNativeRevisionRestoreEvidence,
  type PdfNativeRevisionRestoreEvidence,
} from "./byte-restore-model";
import { matchAllBehind } from "@mrnote/core/web/lib/regexp-behind.ts";

// `xref` at the start or after white space or a delimiter: the lookbehind
// `(?<![^…])`, tested in code (Safari before 16.4 cannot compile one).
const XREF_BEHIND = { char: /[^\x00\t\n\f\r \]>)}]/, not: true } as const;

const ASCII = new TextEncoder();
const LATIN1 = LATIN1_EXACT;

const SHA256_PATTERN = /^[a-f\d]{64}$/;
/** A name as the format writes one: every regular character, which is all of
 * printable ASCII but the delimiters and `#`, whose escapes this engine does
 * not read (PDF 32000-1 7.3.5). A resource this engine only has to recognise
 * -- never to write -- may be named that way; TikZ names every graphics state
 * it declares `pgf@...`. */
const PDF_NAME_TOKEN = /^[\x21-\x22\x24\x26\x27\x2a-\x2e\x30-\x3b\x3d\x3f-\x5a\x5c\x5e-\x7a\x7c\x7e]{1,127}$/;
const MAX_REPLACEMENT_BYTES = 64 * 1024;
// A replacement written as pieces names ranges of the span it replaces
// instead of carrying those bytes again; its own bytes keep the bound above.
// (A format written by state where a paragraph's text stands is a piece or
// two for each run of its shows: a table's cells, a line set glyph by glyph.)
const MAX_REPLACEMENT_PIECES = 8192;
const MAX_DECODED_STREAM_BYTES = 64 * 1024 * 1024;
/** A page's whole drawing, read as one scan. A cover illustration runs to
 * hundreds of thousands of operands (an annual report's front page is two
 * megabytes of paths), and the scan is linear; the decoded stream budget
 * above is what bounds the work. */
const MAX_CONTENT_STREAM_TOKENS = 4_000_000;
/** Width of an end-of-line a producer may have counted in a stream /Length. */
const MAX_FLATE_TRAILING_BYTES = 2;
const MAX_TOP_LEVEL_DICTIONARY_ENTRIES = 4_096;
const MAX_PDF_NAME_BYTES = 256;
const MAX_REFERENCE_COUNT_TARGETS = 16_384;

export function assertMraiNativePageDecodedContentBudget(
  decodedLengths: readonly number[],
): void {
  let total = 0;
  for (const length of decodedLengths) {
    if (!Number.isSafeInteger(length) || length < 0) {
      fail("unsupported-document", "PDF content stream length is invalid.");
    }
    if (length > MAX_DECODED_STREAM_BYTES - total) {
      fail(
        "unsupported-document",
        "Decoded PDF page content exceeds the supported limit.",
      );
    }
    total += length;
  }
}

const MAX_MRAI_NATIVE_EMBEDDED_IMAGE_PIXELS = 40_000_000;

export interface PdfNativeOutlineObjectRef {
  readonly objectNumber: number;
  readonly generation: number;
}

export interface PdfNativeOutlineDestination {
  readonly pageNumber: number;
  readonly page: PdfNativeOutlineObjectRef;
  readonly left: number | null;
  readonly top: number | null;
  readonly zoom: number | null;
}

export interface PdfNativeManagedOutlineNode {
  readonly id: string;
  readonly title: string;
  readonly destination: PdfNativeOutlineDestination;
  readonly children: readonly PdfNativeManagedOutlineNode[];
}

/** A bookmark the file itself carries (not MRAI's own subtree): its object,
 * compared by its current bytes, and for a rename the new title. */
export interface PdfNativeOutlineSourceTarget extends PdfNativeOutlineObjectRef {
  readonly expectedObjectSha256: string;
  readonly title: string | null;
}

export interface PdfNativeOutlineMutationEvidence {
  readonly schemaVersion: typeof MRAI_NATIVE_OUTLINE_SCHEMA_VERSION;
  readonly operation:
    | "add"
    | "rename"
    | "update-destination"
    | "move"
    | "delete"
    | "rename-source"
    | "delete-source";
  readonly expectedCatalog: PdfNativeOutlineObjectRef & {
    readonly expectedObjectSha256: string;
  };
  readonly expectedOutlineRoot:
    | (PdfNativeOutlineObjectRef & { readonly expectedObjectSha256: string })
    | null;
  readonly managedTree: readonly PdfNativeManagedOutlineNode[];
  readonly expectedManagedIds: readonly string[];
  readonly sourceTarget?: PdfNativeOutlineSourceTarget;
  /** The request named the upload's pages on a file holding the page's
   * structure: for each page of the file, the upload's page it is (null for
   * an added one). The tree above is on the file's pages. */
  readonly pageMap?: readonly (number | null)[];
}

export interface PdfNativePreparedOutlineMutation {
  readonly evidence: PdfNativeOutlineMutationEvidence;
  readonly sourceLength: number;
  readonly sourceXrefOffset: number;
  readonly incrementalXrefOffset: number;
  readonly candidateSha256: string;
  readonly candidateBytes: Uint8Array;
  readonly revisedObjects: readonly Readonly<
    PdfNativeOutlineObjectRef & { offset: number; sha256: string }
  >[];
}

export interface PdfNativeOutlineMutationInputNode {
  readonly id: string;
  readonly title: string;
  readonly destination: Readonly<{
    readonly pageNumber: number;
    readonly left: number | null;
    readonly top: number | null;
    readonly zoom: number | null;
  }>;
  readonly children: readonly PdfNativeOutlineMutationInputNode[];
}

export interface PdfNativeOutlineMutationRequest {
  readonly expectedSourceSha256: string;
  readonly operation: PdfNativeOutlineMutationEvidence["operation"];
  readonly expectedManagedIds: readonly string[];
  readonly managedTree: readonly PdfNativeOutlineMutationInputNode[];
  /** rename-source / delete-source: the file's own bookmark by its place in
   * the outline (sibling indexes from the top), and a rename's new title. */
  readonly sourceTarget?: Readonly<{ path: readonly number[]; title: string | null; expectedTitle?: string }>;
  /** The window shows a saved file holding the page's structure under the
   * upload's page numbers (lib/pdf/structured-document): the tree's
   * destinations name the upload's pages, and this is the file's map (for
   * each file page, the upload's page it is; null for an added one). */
  readonly pageMap?: readonly (number | null)[];
}


export interface PdfNativeObjectIdentity {
  readonly sourceSha256: string;
  readonly objectNumber: number;
  readonly generation: number;
}

export interface PdfNativeObjectPrecondition {
  readonly objectNumber: number;
  readonly generation: number;
  readonly expectedObjectSha256: string;
  readonly role: "resource" | "untouched-object";
}

export interface PdfNativeOperatorSpanPrecondition {
  readonly start: number;
  readonly end: number;
  readonly expectedSha256: string;
}

export interface PdfNativeContentStreamInspection {
  readonly identity: PdfNativeObjectIdentity;
  readonly objectSha256: string;
  readonly streamSha256: string;
  readonly streamBytes: Uint8Array;
  readonly untouchedObjects: readonly PdfNativeObjectPrecondition[];
}

export interface PdfNativeObjectInspection {
  readonly streamEncoding?: string | null;
  readonly identity: PdfNativeObjectIdentity;
  readonly objectSha256: string;
  readonly objectBytes: Uint8Array;
  readonly streamSha256: string | null;
  readonly streamBytes: Uint8Array | null;
}

export interface PdfNativePageContentGraphInspection {
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly page: Readonly<{ objectNumber: number; generation: number }>;
  readonly resourceAuthority: Readonly<{
    readonly owner: Readonly<{ objectNumber: number; generation: number }>;
    readonly dictionaryKind: PdfNativeResourceDictionaryKind;
  }>;
  readonly contentStreams: readonly Readonly<{
    objectNumber: number;
    generation: number;
  }>[];
  readonly pageObjectBytes: Uint8Array;
  /**
   * The page's visible box (CropBox, else MediaBox; inherited through the
   * page tree) as `[x0, y0, x1, y1]`, or null when the page declares none the
   * reader could parse. The vector scanners use it to tell a page-sized clip
   * (which restricts nothing on the page) from a real one.
   */
  readonly visibleBox: readonly [number, number, number, number] | null;
  readonly extGStates?: PdfNativePageExtGStateInspection;
  /** Why this page's content could not be read, when the caller asked for
   * the other pages regardless: only the edited page's content has to be
   * readable, and a form shared with such a page stays uncounted. */
  readonly unusableContentStreams?: string;
  /** The refusal itself, for a caller that reads the whole document and
   * fails on the first page it cannot read. */
  readonly unusableContentError?: PdfNativeTransactionError;
}

export interface PdfNativePageExtGStateEntryInspection {
  readonly resourceName: string;
  readonly fillOpacity: number | null;
  readonly strokeOpacity: number | null;
  readonly blendMode: PdfStandardBlendMode | null;
  readonly supportedAppearanceOnly: boolean;
  readonly dictionarySha256: string;
}

export interface PdfNativePageExtGStateInspection {
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly page: Readonly<{
    objectNumber: number;
    generation: number;
    expectedObjectSha256: string;
  }>;
  readonly resourceOwner: Readonly<{
    objectNumber: number;
    generation: number;
    expectedObjectSha256: string;
    dictionaryKind: PdfNativeResourceDictionaryKind;
  }>;
  readonly category: Readonly<
    | { kind: "missing"; expectedDictionarySha256: string }
    | { kind: "direct"; expectedDictionarySha256: string }
    | {
        kind: "indirect";
        expectedDictionarySha256: string;
        target: Readonly<{
          objectNumber: number;
          generation: number;
          expectedObjectSha256: string;
        }>;
      }
  >;
  readonly entries: readonly PdfNativePageExtGStateEntryInspection[];
}

export interface PdfNativePageExtGStateMutationPlan {
  readonly pageNumber: number;
  readonly page: PdfNativePageExtGStateInspection["page"];
  readonly resourceOwner: PdfNativePageExtGStateInspection["resourceOwner"];
  readonly category: PdfNativePageExtGStateInspection["category"];
  readonly resourceName: string;
  readonly fillOpacity: number | null;
  readonly strokeOpacity: number | null;
  readonly blendMode: PdfStandardBlendMode | null;
  readonly reuseExisting: boolean;
}

export interface PdfNativePageFontExtensionMutationPlan {
  readonly pageNumber: number;
  readonly page: PdfNativePageExtGStateInspection["page"];
  readonly resourceOwner: PdfNativePageExtGStateInspection["resourceOwner"];
  readonly category: PdfNativePageExtGStateInspection["category"];
  readonly fontPackage:
    | PdfNativeFontExtensionPackage
    | PdfNativeFontAliasPackage;
  /** Fonts installed with the first one in the same write: each a subset of
   * one slice of a served face (lib/pdf/native-embedded-fonts.ts), its
   * objects numbered on from the end of the package before it. */
  readonly additionalFontPackages?: readonly PdfNativeFontExtensionPackage[];
  /** The font goes into a Form XObject's own resources: the leaf of the
   * form instance isolation chain, whose clone this transaction writes.
   * `page` and `resourceOwner` then name that form. */
  readonly form?: Readonly<{
    objectNumber: number;
    generation: number;
    expectedObjectSha256: string;
  }>;
}

/** Every font package a page font mutation installs, the first one first. */
export function pdfNativePageFontPackages(
  mutation: Pick<PdfNativePageFontExtensionMutationPlan, "fontPackage" | "additionalFontPackages">,
): readonly (PdfNativeFontExtensionPackage | PdfNativeFontAliasPackage)[] {
  return [mutation.fontPackage, ...(mutation.additionalFontPackages ?? [])];
}

/** Every object those packages carry, in object-number order. */
function pageFontPackageObjects(
  mutation: Pick<PdfNativePageFontExtensionMutationPlan, "fontPackage" | "additionalFontPackages">,
) {
  return pdfNativePageFontPackages(mutation).flatMap((fontPackage) => fontPackage.objects);
}


/** The most fonts of the file one write names in a page's resources. */
export const MAX_MRAI_NATIVE_PAGE_FONT_REFERENCES = 4;

/**
 * Fonts the file already holds, named in one page's resources so a write of
 * that page's own text can draw in them: a bold face set again in the regular
 * face of its family, whose subset with the glyph only another page names
 * (pdf-walk-1000 0979 p.1, the title's «بعض»). No font object is written and
 * none is changed. The page's effective Resources are cloned onto the page
 * with one entry more for each font -- a new object the page then names, the
 * way a page-local font extension installs its font -- so a dictionary the
 * page inherits from its tree, or shares with other pages, is never edited
 * in place. Undone, the page's object is written back as it stood.
 */
export interface PdfNativePageFontReferencePlan {
  readonly pageNumber: number;
  readonly page: PdfNativePageExtGStateInspection["page"];
  readonly resourceOwner: PdfNativePageExtGStateInspection["resourceOwner"];
  readonly category: PdfNativePageExtGStateInspection["category"];
  /** The page's new Resources object: the first number the file has free. */
  readonly resources: Readonly<{ objectNumber: number; generation: 0 }>;
  readonly fonts: readonly Readonly<{
    resourceName: string;
    /** A `/Type /Font` dictionary of the file, as it stands. */
    font: Readonly<{ objectNumber: number; generation: number; expectedObjectSha256: string }>;
  }>[];
}

export type PdfNativeContentAppendCategory = "Font" | "XObject" | "ExtGState";
export type PdfNativeContentAppendKind = "text" | "image" | "path";
export type PdfNativeContentAppendDirection = "append" | "remove";

export type PdfNativeContentAppendAddition = Readonly<
  | {
      category: "Font";
      resourceName: string;
      /** Exact newline-free text the package encodes, in Tj order. */
      text: string;
      fontPackage: PdfNativeFontExtensionPackage;
    }
  | {
      category: "XObject";
      resourceName: string;
      reference: Readonly<{ objectNumber: number; generation: 0 }>;
      softMask: Readonly<{ objectNumber: number; generation: 0 }> | null;
      image: PdfNativeEmbeddedImageReplacementData;
    }
  | {
      category: "ExtGState";
      resourceName: string;
      reference: Readonly<{ objectNumber: number; generation: 0 }>;
      fillOpacity: number | null;
      strokeOpacity: number | null;
    }
>;

export interface PdfNativeContentAppendReusedResource {
  readonly category: PdfNativeContentAppendCategory;
  readonly resourceName: string;
  readonly reference: Readonly<{ objectNumber: number; generation: number }>;
  readonly expectedObjectSha256: string;
}

/**
 * One self-contained `q … Q` block appended to (or removed from) the end of a
 * page's single, exclusively owned content stream. Every resource the block
 * names is either installed through a clone-on-write Resources copy or proven
 * to already resolve on the page; the original stream bytes are never edited.
 */
export interface PdfNativeContentAppendPlan {
  readonly schemaVersion: typeof MRAI_NATIVE_CONTENT_APPEND_SCHEMA_VERSION;
  readonly direction: PdfNativeContentAppendDirection;
  readonly appendKind: PdfNativeContentAppendKind;
  readonly pageNumber: number;
  readonly page: PdfNativePageExtGStateInspection["page"];
  readonly resourceOwner: PdfNativePageExtGStateInspection["resourceOwner"];
  readonly categories: Readonly<
    Partial<
      Record<
        PdfNativeContentAppendCategory,
        PdfNativePageExtGStateInspection["category"]
      >
    >
  >;
  readonly additions: readonly PdfNativeContentAppendAddition[];
  readonly reusedResources: readonly PdfNativeContentAppendReusedResource[];
  /** Unbalanced `q` levels the source stream leaves open; the block pops them first. */
  readonly restoredDepth: number;
}

export interface PdfNativePreparedContentAppend {
  readonly objects: readonly PdfNativePreparedIsolatedObject[];
  readonly resources: PdfNativePreparedIsolatedObject | null;
  readonly page: PdfNativePreparedIsolatedObject | null;
}

export interface PdfNativeContentAppendCategoryInspection {
  readonly state: PdfNativePageExtGStateInspection["category"];
  readonly entries: readonly PdfNativeReferenceDictionaryEntry[];
}

export interface PdfNativeContentAppendPageInspection {
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly page: PdfNativePageExtGStateInspection["page"];
  readonly resourceOwner: PdfNativePageExtGStateInspection["resourceOwner"];
  readonly contentStream: PdfNativeContentStreamInspection;
  /**
   * The rest of the page drawing, ahead of the stream the block is appended
   * to. A page drawn from several streams leaves its tail graphics state
   * across all of them, so the block's restore prefix is read from the whole.
   */
  readonly drawingPrefix: Uint8Array;
  readonly categories: Readonly<
    Record<
      PdfNativeContentAppendCategory,
      PdfNativeContentAppendCategoryInspection
    >
  >;
  readonly firstNewObjectNumber: number;
}

export const MRAI_NATIVE_FONT_ALIAS_PACKAGE_SCHEMA_VERSION =
  "mrai.native-font-alias-package.v1" as const;

export interface PdfNativeFontAliasPackage {
  readonly schemaVersion: typeof MRAI_NATIVE_FONT_ALIAS_PACKAGE_SCHEMA_VERSION;
  readonly fontSha256: typeof PDF_UNICODE_FALLBACK_FONT_SHA256;
  readonly resourceName: string;
  readonly font: Readonly<{ objectNumber: number; generation: 0 }>;
  readonly sourceFont: Readonly<{
    objectNumber: number;
    generation: number;
    expectedObjectSha256: string;
  }>;
  readonly sourceToUnicode: Readonly<{
    objectNumber: number;
    generation: number;
    expectedObjectSha256: string;
  }>;
  readonly encodedHex: string;
  readonly widthAt1000: number;
  readonly objectByteLength: number;
  readonly objects: readonly Readonly<{
    objectNumber: number;
    generation: 0;
    sha256: string;
    bytes: Uint8Array;
  }>[];
}

/**
 * One piece of a replacement written as pieces: a range of the span it takes
 * the place of, or bytes of its own.
 */
export type PdfNativeReplacementPiece =
  | Readonly<{ spanStart: number; spanEnd: number }>
  | Readonly<{ literal: Uint8Array }>;

export interface PdfNativeContentStreamTransactionRequest {
  readonly identity: PdfNativeObjectIdentity;
  readonly expectedObjectSha256: string;
  readonly expectedStreamSha256: string;
  readonly operatorSpan: PdfNativeOperatorSpanPrecondition;
  readonly replacement: Uint8Array;
  /**
   * The replacement as ranges of the span and the bytes it adds. A picture
   * move rewrites a span it only reorders, so its record names those bytes
   * by where they are and a large drawing moves without being carried again.
   */
  readonly replacementPieces?: readonly PdfNativeReplacementPiece[];
  /**
   * Other spans of the stream that hold the same show, byte for byte, and
   * take the same replacement: a page that draws a line twice in one place
   * (a printer driver's banded output, a fake bold) is edited in every copy,
   * so no copy keeps showing the old text under the new. Each lies outside
   * the operator span and the others, in the order they are written.
   */
  readonly duplicateSpans?: readonly PdfNativeDuplicateSpan[];
  /**
   * The underline and strikethrough drawn beside the edited run's text
   * object, rewritten in this same write (a property edit that moves the
   * run): each names the rule bytes now standing at its span, an underline
   * ending at the object's `BT`, a strikethrough starting after its `ET`,
   * and the bytes written there instead. One write, one undo.
   */
  readonly textRules?: readonly PdfNativeTextRuleSpan[];
  /**
   * The box this write sets, wrapped in its own `/MRAIBox BMC ... EMC` in
   * the same write (box-mark.ts): where the two halves stand in the source
   * stream, round the operator span. Only a plain write of one span takes
   * one. Optional: a record written before it replays unchanged.
   */
  readonly boxMark?: PdfNativeBoxMark;
  readonly untouchedObjects?: readonly PdfNativeObjectPrecondition[];
  /**
   * Source-derived clone-on-write evidence for one selected Form XObject
   * instance. The stable transaction persists this bounded identity chain,
   * while the server independently reconstructs and compares it against the
   * pinned bytes before admitting replay.
   */
  readonly formInstanceIsolation?: PdfNativeFormInstanceIsolationPlan;
  /**
   * One page-local graphics-state addition. The effective Resources tree is
   * cloned onto the selected page; inherited or shared dictionaries are never
   * edited in place.
   */
  readonly pageExtGStateMutation?: PdfNativePageExtGStateMutationPlan;
  /**
   * One `/XObject` name dropped from (or put back into) the page's own
   * resource dictionary, so a deleted picture stops being reachable and a
   * later full rewrite can drop its bytes.
   */
  readonly pageXObjectUnmap?: PdfNativePageXObjectUnmapPlan;
  /** One verified Type0 subset installed through page-local Resources COW. */
  readonly pageFontExtensionMutation?: PdfNativePageFontExtensionMutationPlan;
  /** Fonts the file already holds, named in the page's resources through
   * the same clone-on-write; rides one plain write of the page's own text. */
  readonly pageFontReference?: PdfNativePageFontReferencePlan;
  /**
   * One self-contained block appended to, or removed from, the tail of the
   * page's single content stream, with its resource installations.
   */
  readonly contentAppend?: PdfNativeContentAppendPlan;
  /**
   * A page whose content is split across several streams at places no
   * stream can be read on its own (a text object or an operator's operands
   * straddle a boundary) is edited as the one stream a reader sees: its
   * decoded streams joined, written as one new object the page's
   * `/Contents` then names. The original streams stay in the file untouched.
   */
  readonly pageContentConsolidation?: PdfNativePageContentConsolidationPlan;
  




  readonly structActualText?: readonly PdfNativeStructActualTextRevision[];
}

export interface PdfNativePageContentConsolidationPlan {
  readonly pageNumber: number;
  readonly page: Readonly<{
    objectNumber: number;
    generation: number;
    expectedObjectSha256: string;
  }>;
  readonly sourceStreams: readonly Readonly<{
    objectNumber: number;
    generation: number;
    expectedObjectSha256: string;
  }>[];
  readonly joinedDecodedSha256: string;
  readonly stream: Readonly<{ objectNumber: number; generation: 0 }>;
}

export interface PdfNativePreparedPageContentConsolidation {
  readonly stream: PdfNativePreparedIsolatedObject;
  readonly page: PdfNativePreparedIsolatedObject;
}

export interface PdfNativeEmbeddedImageRemovalEvidence {
  readonly schemaVersion: typeof MRAI_NATIVE_EMBEDDED_IMAGE_REMOVE_SCHEMA_VERSION;
  readonly pageNumber: number;
  readonly imageOrder: number;
  readonly resourceName: string;
  readonly page: Readonly<{ objectNumber: number; generation: number }>;
  readonly resourceOwner: Readonly<{
    objectNumber: number;
    generation: number;
  }>;
  readonly image: Readonly<{ objectNumber: number; generation: number }>;
  readonly contentStream: Readonly<{
    objectNumber: number;
    generation: number;
  }>;
  readonly operatorSpan: Readonly<{ start: number; end: number }>;
  readonly drawOperatorBytes: readonly number[];
  readonly removedOperatorBytes: readonly number[];
  /**
   * Present when this revision also drops the picture's `/XObject` name from
   * the page's own resources, which is what lets a later full rewrite drop
   * the image bytes. The server recomputes the rewritten owner itself.
   */
  readonly resourceUnmap?: Readonly<{
    resourceName: string;
    direction: "unmap" | "remap";
    resourceOwner: Readonly<{ objectNumber: number; generation: number }>;
    expectedOwnerSha256: string;
    revisedOwnerSha256: string;
    entryOffset: number;
    entryBytes: readonly number[];
  }>;
}

export interface PdfNativeEmbeddedImageRemovalPlan {
  readonly evidence: PdfNativeEmbeddedImageRemovalEvidence;
  readonly request: PdfNativeContentStreamTransactionRequest;
}

export interface PdfNativeEmbeddedImageRemovalSelection {
  readonly pageNumber: number;
  readonly imageOrder: number;
  /** The CTM the reader measured the picture at: when given, the paint at
   * `imageOrder` must be drawn there (see `resolvePdfNativeEmbeddedImageOrder`). */
  readonly matrix?: PdfNativeEmbeddedImageTransformMatrix | null;
}

export type PdfNativeEmbeddedImageObjectRef = Readonly<{
  objectNumber: number;
  generation: number;
}>;

/** `[a b c d e f]` written before the image draw, concatenated onto the CTM. */
export type PdfNativeEmbeddedImageTransformMatrix = readonly [
  number,
  number,
  number,
  number,
  number,
  number,
];

/** A clip in the image's own unit square (bottom-left origin, 0..1). */
export interface PdfNativeEmbeddedImageTransformClip {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

/**
 * One source image draw rewritten in place as `q <matrix> cm [<clip> re W n]
 * /<name> Do Q`. The image object, its resource mapping and every other
 * operator stay byte-exact; only the selected paint instance moves, scales,
 * rotates, flips or clips. Undo/redo swap the two operator byte strings.
 */
export interface PdfNativeEmbeddedImageTransformEvidence {
  readonly schemaVersion: typeof MRAI_NATIVE_EMBEDDED_IMAGE_TRANSFORM_SCHEMA_VERSION;
  readonly pageNumber: number;
  readonly imageOrder: number;
  readonly resourceName: string;
  readonly page: PdfNativeEmbeddedImageObjectRef;
  readonly resourceOwner: PdfNativeEmbeddedImageObjectRef;
  readonly image: PdfNativeEmbeddedImageObjectRef;
  readonly contentStream: PdfNativeEmbeddedImageObjectRef;
  readonly originalOperatorBytes: readonly number[];
  readonly transformedOperatorBytes: readonly number[];
  readonly matrix: PdfNativeEmbeddedImageTransformMatrix;
  readonly clip: PdfNativeEmbeddedImageTransformClip | null;
  /** ExtGState resource named in the wrapper, when this paint carries alpha. */
  readonly graphicsStateName?: string | null;
  readonly opacity?: number | null;
}

export interface PdfNativeEmbeddedImageTransformPlan {
  readonly evidence: PdfNativeEmbeddedImageTransformEvidence;
  readonly request: PdfNativeContentStreamTransactionRequest;
}

export interface PdfNativeEmbeddedImageTransformSelection {
  readonly pageNumber: number;
  readonly imageOrder: number;
}

export interface PdfNativeEmbeddedImageReplacementData {
  readonly mimeType: "image/jpeg" | "image/png";
  readonly width: number;
  readonly height: number;
  readonly colorSpace: "DeviceGray" | "DeviceRGB";
  readonly bitsPerComponent: 8;
  readonly filter: "DCTDecode" | "FlateDecode";
  readonly decodeParms: Readonly<{
    predictor: 15;
    colors: 1 | 3;
    bitsPerComponent: 8;
    columns: number;
  }> | null;
  /** Encoded JPEG bytes or the concatenated PNG IDAT payload. */
  readonly encodedDataBase64: string;
  readonly encodedSha256: string;
  /** Same-sized 8-bit grayscale PNG-predictor opacity plane, never a flatten. */
  readonly softMask?: Readonly<{
    encodedDataBase64: string;
    encodedSha256: string;
  }>;
}

export interface PdfNativeEmbeddedImageReplacementEvidence {
  readonly schemaVersion: typeof MRAI_NATIVE_EMBEDDED_IMAGE_REPLACE_SCHEMA_VERSION;
  readonly pageNumber: number;
  readonly imageOrder: number;
  readonly originalResourceName: string;
  readonly replacementResourceName: string;
  readonly page: Readonly<{ objectNumber: number; generation: number }>;
  readonly resourceOwner: Readonly<{
    objectNumber: number;
    generation: number;
  }>;
  readonly originalImage: Readonly<{
    objectNumber: number;
    generation: number;
  }>;
  readonly replacementImage: Readonly<{
    objectNumber: number;
    generation: number;
    expectedObjectSha256: string;
  }>;
  readonly replacementSoftMask?: Readonly<{
    objectNumber: number;
    generation: number;
    expectedObjectSha256: string;
  }>;
  readonly contentStream: Readonly<{
    objectNumber: number;
    generation: number;
  }>;
  readonly originalOperatorBytes: readonly number[];
  readonly replacementOperatorBytes: readonly number[];
  readonly replacement: PdfNativeEmbeddedImageReplacementData;
}

export interface PdfNativePreparedEmbeddedImageReplacement {
  readonly evidence: PdfNativeEmbeddedImageReplacementEvidence;
  readonly sourceLength: number;
  readonly sourceXrefOffset: number;
  readonly incrementalXrefOffset: number;
  readonly candidateSha256: string;
  readonly candidateBytes: Uint8Array;
  readonly revisedObjects: readonly Readonly<{
    objectNumber: number;
    generation: number;
    offset: number;
    sha256: string;
  }>[];
  /** The catalog's XMP packet written last without its PDF/A
   * identification (also one of `revisedObjects`). */
  readonly pdfaIdentificationRemoval?: PdfNativePreparedPdfaIdentificationRemoval;
  /** The catalog stating the version the revision needs (PDF 0020). */
  readonly catalogVersionRaise?: PdfNativePreparedCatalogVersionRaise;
}

export interface PdfNativeEmbeddedImageReplacementSelection {
  readonly pageNumber: number;
  readonly imageOrder: number;
  /** As the removal's: the CTM the reader measured the picture at. */
  readonly matrix?: PdfNativeEmbeddedImageTransformMatrix | null;
}

/** One drawing of a Form its parent draws more than once: that one `Do` is
 * renamed, and the parent's resources map the new name to the Form's clone,
 * so the edit reaches that drawing alone and every other drawing keeps the
 * original (mirror of the server's NativeFormIsolationInstance). */
export interface PdfNativeFormInstanceIsolationInstance {
  /** The stream holding the `Do`: one of the page's content streams for the
   * first Form of the chain, the Form before it otherwise. */
  readonly contentStream: Readonly<{ objectNumber: number; generation: number }>;
  /** Offset of the `Do`'s name operand in that stream's decoded bytes. */
  readonly operandStart: number;
  /** The name the parent's resources give the clone. */
  readonly resourceName: string;
}

export interface PdfNativeFormInstanceIsolationPlan {
  readonly page: Readonly<{
    objectNumber: number;
    generation: number;
    expectedObjectSha256: string;
  }>;
  /** Top-level page resource first, selected leaf Form last. A link is
   * named by its parent either as an /XObject resource (the default) or, for
   * a resources dictionary the pages share, as the parent's own /Resources:
   * that dictionary is cloned too, so the edited page stops sharing it. A
   * Form its parent draws more than once carries the one drawing edited. */
  readonly chain: readonly Readonly<{
    resourceName: string;
    kind?: "xobject" | "resources";
    objectNumber: number;
    generation: number;
    expectedObjectSha256: string;
    instance?: PdfNativeFormInstanceIsolationInstance;
  }>[];
}

interface PdfNativePreparedIsolatedObject {
  readonly objectNumber: number;
  readonly generation: number;
  readonly offset: number;
  readonly sha256: string;
}

interface PdfNativePreparedFormIsolation {
  readonly primaryClone: Readonly<{
    objectNumber: number;
    generation: number;
  }>;
  readonly objects: readonly PdfNativePreparedIsolatedObject[];
  /** Each link of the chain and its copy, leaf first: what a caller reads
   * the candidate's proofs back through. */
  readonly clones?: readonly Readonly<{
    original: Readonly<{ objectNumber: number; generation: number }>;
    clone: Readonly<{ objectNumber: number; generation: number }>;
  }>[];
}

export type PdfNativePreparedPageExtGStateMutation = Readonly<
  | { resourceName: string; reused: true }
  | {
      resourceName: string;
      reused: false;
      graphicsState: Readonly<{
        objectNumber: number;
        generation: number;
        offset: number;
        sha256: string;
      }>;
      resources: Readonly<{
        objectNumber: number;
        generation: number;
        offset: number;
        sha256: string;
      }>;
      page: Readonly<{
        objectNumber: number;
        generation: number;
        offset: number;
        sha256: string;
      }>;
    }
>;

export interface PdfNativePreparedPageFontExtensionMutation {
  readonly resourceName: string;
  readonly font: Readonly<{ objectNumber: number; generation: 0 }>;
  readonly fontObjects: readonly PdfNativePreparedIsolatedObject[];
  readonly resources: PdfNativePreparedIsolatedObject;
  readonly page: PdfNativePreparedIsolatedObject;
}

/** Where a page font reference's two objects stand in the candidate. */
export interface PdfNativePreparedPageFontReference {
  readonly fonts: readonly Readonly<{
    resourceName: string;
    font: Readonly<{ objectNumber: number; generation: number }>;
  }>[];
  readonly resources: PdfNativePreparedIsolatedObject;
  readonly page: PdfNativePreparedIsolatedObject;
}

export interface PdfNativeReusablePageFontExtension {
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly resourceName: string;
  readonly font: Readonly<{ objectNumber: number; generation: 0 }>;
  /** The source font dictionary an alias package is cloned from. */
  readonly fontObject: Readonly<{
    objectNumber: number;
    generation: number;
    objectSha256: string;
    objectBytes: Uint8Array;
  }>;
  readonly resources: Readonly<{
    page: PdfNativeIndirectObjectRef;
    font: PdfNativeIndirectObjectRef;
    descendantFont: PdfNativeIndirectObjectRef;
    descriptor: PdfNativeIndirectObjectRef;
    fontFile: PdfNativeIndirectObjectRef;
    toUnicode: PdfNativeIndirectObjectRef;
  }>;
  readonly fontPackage: PdfNativeFontExtensionPackage;
  readonly sharedObjects: readonly Readonly<{
    objectNumber: number;
    generation: 0;
    sha256: string;
  }>[];
  readonly currentToUnicode: Readonly<{
    objectNumber: number;
    generation: number;
    objectSha256: string;
    streamSha256: string;
    streamBytes: Uint8Array;
    objectBytes: Uint8Array;
  }>;
  readonly replacementToUnicode: Readonly<{
    objectNumber: number;
    generation: 0;
    objectSha256: string;
    streamSha256: string;
    streamBytes: Uint8Array;
    objectBytes: Uint8Array;
  }>;
}

export interface PdfNativeFontExtensionHistoryObjectState {
  readonly targetObjectSha256: string;
  readonly targetStreamSha256: string;
  readonly pageObjectSha256: string;
}

export interface PdfNativePreparedFontExtensionHistory {
  readonly direction: "undo" | "redo";
  readonly target: PdfNativeObjectIdentity;
  readonly page: Readonly<{ objectNumber: number; generation: number }>;
  readonly current: PdfNativeFontExtensionHistoryObjectState;
  readonly desired: PdfNativeFontExtensionHistoryObjectState;
  readonly sourceLength: number;
  readonly sourceXrefOffset: number;
  readonly incrementalXrefOffset: number;
  readonly changedObjectOffset: number;
  readonly changedObjectSha256: string;
  readonly changedStreamSha256: string;
  readonly candidateSha256: string;
  readonly candidateBytes: Uint8Array;
  readonly revisedObjects: readonly Readonly<{
    objectNumber: number;
    generation: number;
    offset: number;
    sha256: string;
  }>[];
}

export interface PdfNativePreparedContentStreamTransaction {
  readonly schemaVersion: typeof MRAI_NATIVE_CONTENT_STREAM_SCHEMA_VERSION;
  readonly request: PdfNativeContentStreamTransactionRequest;
  readonly sourceLength: number;
  readonly sourceXrefOffset: number;
  /** How many xref sections the source's chain holds, as prepare read it:
   * what a checkpoint claim compares, without parsing the source again. */
  readonly sourceXrefSections?: number;
  readonly changedObjectOffset: number;
  readonly incrementalXrefOffset: number;
  readonly changedObjectSha256: string;
  readonly changedStreamSha256: string;
  readonly changedDecodedStreamSha256: string;
  readonly replacementSha256: string;
  readonly candidateSha256: string;
  readonly candidateBytes: Uint8Array;
  readonly formInstanceIsolation?: PdfNativePreparedFormIsolation;
  readonly pageExtGStateMutation?: PdfNativePreparedPageExtGStateMutation;
  readonly pageXObjectUnmap?: Readonly<{
    resourceName: string;
    direction: "unmap" | "remap";
    resourceOwner: PdfNativePreparedIsolatedObject;
  }>;
  readonly pageFontExtensionMutation?: PdfNativePreparedPageFontExtensionMutation;
  readonly pageFontReference?: PdfNativePreparedPageFontReference;
  readonly contentAppend?: PdfNativePreparedContentAppend;
  readonly pageContentConsolidation?: PdfNativePreparedPageContentConsolidation;
  /** Each structure element the revision writes after the content write's
   * own objects, where it stands and its digest. */
  readonly structActualText?: readonly PdfNativePreparedIsolatedObject[];
  

  readonly pdfaIdentificationRemoval?: PdfNativePreparedPdfaIdentificationRemoval;
  /** The catalog stating the version the revision needs (PDF 0020). */
  readonly catalogVersionRaise?: PdfNativePreparedCatalogVersionRaise;
}

export interface PdfNativeReopenVerification {
  readonly verified: true;
  readonly identity: PdfNativeObjectIdentity;
  readonly sourceXrefOffset: number;
  readonly incrementalXrefOffset: number;
  readonly candidateSha256: string;
}

interface XrefEntry {
  readonly objectNumber: number;
  readonly generation: number;
  readonly offset: number;
  readonly inUse: boolean;
  readonly kind: "free" | "direct" | "compressed";
  readonly objectStreamNumber: number | null;
  readonly objectStreamIndex: number | null;
}

interface PdfXrefSection {
  readonly kind: "classic" | "stream";
  readonly offset: number;
  readonly entries: ReadonlyMap<number, XrefEntry>;
  readonly trailerRaw: string;
  readonly size: number;
  readonly rootObjectNumber: number;
  readonly rootGeneration: number;
  readonly previousOffset: number | null;
  /** On the newest section only: every object whose newest entry in the
   * chain misses it, at the place the file was found to hold it. A revision
   * appended over this one writes these rows again where they really are. */
  readonly relocated?: ReadonlyMap<number, XrefEntry>;
}

interface ParsedPdf {
  readonly current: PdfXrefSection;
  readonly previous: PdfXrefSection | null;
  readonly sections: readonly PdfXrefSection[];
}

interface ParsedIndirectObject {
  readonly objectNumber: number;
  readonly generation: number;
  readonly storage: "direct" | "object-stream";
  readonly start: number;
  readonly end: number;
  readonly bytes: Uint8Array;
  readonly dictionaryStart: number | null;
  readonly dictionaryEnd: number | null;
  readonly lengthNumberStart: number | null;
  readonly lengthNumberEnd: number | null;
  /** The stream dictionary names no /Length at all; readers find the end at
   * its one `endstream`. The length span is then empty, just inside `<<`, and
   * a revision writes the whole `/Length N ` there. */
  readonly lengthKeyMissing?: boolean;
  /** Set when the producer stored /Length in its own object instead of
   * inline; there are then no local digits to patch. */
  readonly lengthReference: Readonly<{
    objectNumber: number;
    generation: number;
  }> | null;
  readonly streamStart: number | null;
  readonly streamEnd: number | null;
  readonly streamEncoding:
    | "identity"
    | "flate"
    | "flate-predictor15"
    | "flate-png-predictor"
    | "ascii85-flate"
    | "ascii-hex"
    | "lzw"
    | "dct"
    | "ascii85-dct"
    /** Read for its producer bytes only (an object pinned untouched is
     * hashed, never decoded); `streamBytes` is null. */
    | "opaque"
    | null;
  readonly encodedStreamBytes: Uint8Array | null;
  readonly streamBytes: Uint8Array | null;
}


function bytesEqual(left: Uint8Array, right: Uint8Array): boolean {
  if (left.length !== right.length) return false;
  let index = 0;
  // Four bytes at a time where both buffers allow it: a save compares its
  // whole source with the candidate's prefix.
  if (left.byteOffset % 4 === 0 && right.byteOffset % 4 === 0) {
    const words = left.length >>> 2;
    const leftWords = new Uint32Array(left.buffer, left.byteOffset, words);
    const rightWords = new Uint32Array(right.buffer, right.byteOffset, words);
    for (; index < words; index += 1) {
      if (leftWords[index] !== rightWords[index]) return false;
    }
    index = words << 2;
  }
  for (; index < left.length; index += 1) {
    if (left[index] !== right[index]) return false;
  }
  return true;
}

function isUint8Array(value: unknown): value is Uint8Array {
  return Object.prototype.toString.call(value) === "[object Uint8Array]";
}

function concatBytes(parts: readonly Uint8Array[]): Uint8Array {
  const total = parts.reduce((sum, part) => sum + part.length, 0);
  const joined = new Uint8Array(total);
  let cursor = 0;
  for (const part of parts) {
    joined.set(part, cursor);
    cursor += part.length;
  }
  return joined;
}

function imagePredictorParameters(dictionary: string): {
  columns: number;
  colors: number;
  bits: number;
} | null {
  if (!dictionary.startsWith("<<") || !dictionary.endsWith(">>")) return null;
  const entries = topLevelDictionaryEntries(dictionary);
  const allowed = new Set([
    "Predictor",
    "Colors",
    "BitsPerComponent",
    "Columns",
  ]);
  if (entries.some((entry) => !allowed.has(entry.key))) return null;
  const number = (key: string, fallback: number) =>
    entries.some((entry) => entry.key === key)
      ? oneTopLevelInteger(dictionary, key, `PDF image predictor ${key}`)
      : fallback;
  const predictor = number("Predictor", 1);
  const colors = number("Colors", 1);
  const bits = number("BitsPerComponent", 8);
  const columns = number("Columns", 1);
  return predictor === 15 &&
    (colors === 1 || colors === 3) &&
    bits === 8 &&
    columns > 0
    ? { columns, colors, bits }
    : null;
}

/** PNG-predicted cross-reference stream rows: one filter byte per row,
 * eight-bit single-component samples, Columns equal to the /W row width. */
function xrefPredictorParameters(dictionary: string): { columns: number } | null {
  if (!dictionary.startsWith("<<") || !dictionary.endsWith(">>")) return null;
  const entries = topLevelDictionaryEntries(dictionary);
  const allowed = new Set(["Predictor", "Colors", "BitsPerComponent", "Columns"]);
  if (entries.some((entry) => !allowed.has(entry.key))) return null;
  const number = (key: string, fallback: number) =>
    entries.some((entry) => entry.key === key)
      ? oneTopLevelInteger(dictionary, key, `PDF xref predictor ${key}`)
      : fallback;
  const predictor = number("Predictor", 1);
  const colors = number("Colors", 1);
  const bits = number("BitsPerComponent", 8);
  const columns = number("Columns", 1);
  return predictor >= 10 && predictor <= 15 && colors === 1 && bits === 8 && columns >= 1 && columns <= 64
    ? { columns }
    : null;
}

function unpredictPngRows(data: Uint8Array, columns: number): Uint8Array {
  const rowLength = columns + 1;
  if (data.length === 0 || data.length % rowLength !== 0) {
    fail("unsupported-document", "PDF xref stream predictor rows are inconsistent.");
  }
  const rows = data.length / rowLength;
  const output = new Uint8Array(rows * columns);
  const previous = new Uint8Array(columns);
  for (let row = 0; row < rows; row += 1) {
    const filter = data[row * rowLength];
    const input = data.subarray(row * rowLength + 1, (row + 1) * rowLength);
    const current = output.subarray(row * columns, (row + 1) * columns);
    for (let index = 0; index < columns; index += 1) {
      const raw = input[index];
      const left = index > 0 ? current[index - 1] : 0;
      const up = previous[index];
      const upLeft = index > 0 ? previous[index - 1] : 0;
      let value: number;
      if (filter === 0) value = raw;
      else if (filter === 1) value = raw + left;
      else if (filter === 2) value = raw + up;
      else if (filter === 3) value = raw + Math.floor((left + up) / 2);
      else if (filter === 4) {
        const estimate = left + up - upLeft;
        const distanceLeft = Math.abs(estimate - left);
        const distanceUp = Math.abs(estimate - up);
        const distanceUpLeft = Math.abs(estimate - upLeft);
        value =
          raw +
          (distanceLeft <= distanceUp && distanceLeft <= distanceUpLeft
            ? left
            : distanceUp <= distanceUpLeft
              ? up
              : upLeft);
      } else {
        fail("unsupported-document", "PDF xref stream row filter is invalid.");
      }
      current[index] = value & 0xff;
    }
    previous.set(current);
  }
  return output;
}

/**
 * Which filter chain a stream carries, and so how its bytes are read.
 *
 * `allowOpaque` is for the readers that never look at the bytes -- proving an
 * object is unchanged is hashing what the producer wrote, not decoding it.
 * They get `"opaque"` where this engine cannot decode the chain (JPX, CCITT,
 * JBIG2, a predictor it does not read) instead of a refusal, and the decoded
 * bytes are then simply absent. Every writer goes through
 * `assertTargetStream`, which takes none of them.
 */
function parseStreamEncoding(
  dictionary: string,
  allowOpaque = false,
):
  | "identity"
  | "flate"
  | "flate-predictor15"
  | "flate-png-predictor"
  | "ascii85-flate"
  | "ascii-hex"
  | "lzw"
  | "dct"
  | "ascii85-dct"
  | "opaque" {
  // A reader that never looks at the bytes may carry the stream opaquely;
  // otherwise the caller's own reason is why this stream cannot be read. It
  // is never the XFA reason, which belongs to a document that declares one.
  const undecodable = <T>(refuse: () => T): T | "opaque" =>
    allowOpaque ? "opaque" : refuse();
  const entries = topLevelDictionaryEntries(dictionary);
  const filterEntries = entries.filter((entry) => entry.key === "Filter");
  const decodeParmsEntries = entries.filter(
    (entry) => entry.key === "DecodeParms",
  );
  if (filterEntries.length === 0) {
    if (decodeParmsEntries.length !== 0) {
      return undecodable(() =>
        fail(
          "unsupported-document",
          "PDF stream DecodeParms require one supported Flate filter.",
        ),
      );
    }
    return "identity";
  }
  if (filterEntries.length !== 1) {
    return undecodable(() =>
      fail("unsupported-document", "PDF stream filters are ambiguous."),
    );
  }
  const filter = dictionary.slice(
    filterEntries[0].valueStart,
    filterEntries[0].valueEnd,
  );
  const flate =
    /^(?:\/(?:FlateDecode|Fl)|\[\s*\/(?:FlateDecode|Fl)\s*\])$/.test(filter);
  const ascii85Flate =
    /^\[\s*\/(?:ASCII85Decode|A85)\s+\/(?:FlateDecode|Fl)\s*\]$/.test(filter);
  const asciiHex = /^(?:\/(?:ASCIIHexDecode|AHx)|\[\s*\/(?:ASCIIHexDecode|AHx)\s*\])$/.test(filter);
  const lzw = /^(?:\/(?:LZWDecode|LZW)|\[\s*\/(?:LZWDecode|LZW)\s*\])$/.test(filter);
  const dct = /^\/(?:DCTDecode|DCT)$/.test(filter);
  const ascii85Dct =
    /^\[\s*\/(?:ASCII85Decode|A85)\s+\/(?:DCTDecode|DCT)\s*\]$/.test(filter);
  if (!flate && !ascii85Flate && !asciiHex && !lzw && !dct && !ascii85Dct) {
    return undecodable(() =>
      fail(
        "unsupported-document",
        "Only bounded FlateDecode or DCTDecode filter chains are supported.",
      ),
    );
  }
  if (decodeParmsEntries.length > 1) {
    return undecodable(() =>
      fail("unsupported-document", "PDF stream DecodeParms are ambiguous."),
    );
  }
  if (decodeParmsEntries.length === 1) {
    const decodeParms = dictionary.slice(
      decodeParmsEntries[0].valueStart,
      decodeParmsEntries[0].valueEnd,
    );
    // Parameters are neutral when they declare no predictor, or predictor 1:
    // `<<>>`, `null`, `[null]`, and a dictionary that only describes the
    // image's own layout (BitsPerComponent, Colors, Columns) all decode as
    // plain Flate.
    const neutral = (() => {
      const inner = /^\[\s*([\s\S]*?)\s*\]$/.exec(decodeParms)?.[1] ?? decodeParms;
      if (/^(?:null)?$/.test(inner.trim())) return true;
      if (!/^<<[\s\S]*>>$/.test(inner.trim())) return false;
      const predictor = /\/Predictor\s+(\d+)/.exec(inner);
      if (predictor && predictor[1] !== "1") return false;
      const residue = inner
        .trim()
        .slice(2, -2)
        .replace(/\/(?:Predictor|BitsPerComponent|Colors|Columns)\s+\d+/g, "")
        .trim();
      return residue === "";
    })();
    const imagePredictor = imagePredictorParameters(decodeParms);
    const xrefPredictor =
      !neutral && !imagePredictor && flate && /\/Type\s*\/XRef\b/.test(dictionary)
        ? xrefPredictorParameters(decodeParms)
        : null;
    // An LZW chain is read and rewritten with EarlyChange 1, the default; a
    // producer that declares any DecodeParms at all may be asking for
    // EarlyChange 0 or a predictor, neither of which this codec reproduces.
    if (ascii85Flate || asciiHex || lzw || dct || ascii85Dct || (!neutral && !imagePredictor && !xrefPredictor)) {
      return undecodable(() =>
        fail(
          "unsupported-document",
          "Only neutral or bounded image Flate DecodeParms are supported.",
        ),
      );
    }
    if (imagePredictor) return "flate-predictor15";
    if (xrefPredictor) return "flate-png-predictor";
  }
  return ascii85Dct
    ? "ascii85-dct"
    : dct
      ? "dct"
      : asciiHex
        ? "ascii-hex"
        : lzw
          ? "lzw"
          : ascii85Flate
            ? "ascii85-flate"
            : "flate";
}

function isPdfWhitespaceByte(value: number): boolean {
  return (
    value === 0x00 ||
    value === 0x09 ||
    value === 0x0a ||
    value === 0x0c ||
    value === 0x0d ||
    value === 0x20
  );
}

/** ASCIIHexDecode: each byte written as two hex digits, whitespace ignored,
 * `>` ends the data and an odd last digit is padded with zero (PDF 32000-1
 * 7.4.2). */
function decodeAsciiHexStream(encoded: Uint8Array): Uint8Array {
  const decoded: number[] = [];
  let high: number | null = null;
  for (const value of encoded) {
    if (value === 0x3e) break;
    if (isPdfWhitespaceByte(value)) continue;
    const digit =
      value >= 0x30 && value <= 0x39
        ? value - 0x30
        : value >= 0x41 && value <= 0x46
          ? value - 0x37
          : value >= 0x61 && value <= 0x66
            ? value - 0x57
            : -1;
    if (digit < 0) {
      fail("unsupported-document", "ASCIIHex stream holds a non-hex byte.");
    }
    if (high === null) {
      high = digit;
      continue;
    }
    decoded.push(high * 16 + digit);
    high = null;
    if (decoded.length > MAX_DECODED_STREAM_BYTES) {
      fail("unsupported-document", "ASCIIHex stream is too large.");
    }
  }
  if (high !== null) decoded.push(high * 16);
  return Uint8Array.from(decoded);
}

function encodeAsciiHexStream(decoded: Uint8Array): Uint8Array {
  const digits = "0123456789ABCDEF";
  const encoded: number[] = [];
  for (const value of decoded) {
    encoded.push(digits.charCodeAt(value >>> 4), digits.charCodeAt(value & 0x0f));
  }
  encoded.push(0x3e);
  return Uint8Array.from(encoded);
}

/** LZW (PDF 32000-1 7.4.4) with the default EarlyChange 1: codes are packed
 * most-significant bit first, 256 clears the table, 257 ends the data, and
 * the first assigned code is 258. The code width grows from 9 to 12 one
 * assignment before the table would need it, which is what EarlyChange 1
 * means; a stream that declares DecodeParms never reaches this codec, so
 * EarlyChange 0 and predictored chains stay refused rather than decoded
 * wrongly. */
const LZW_CLEAR_CODE = 256;
const LZW_END_CODE = 257;
const LZW_FIRST_CODE = 258;
const LZW_MAX_CODE = 4096;

function decodeLzwStream(encoded: Uint8Array): Uint8Array {
  const decoded: number[] = [];
  const table: Uint8Array[] = [];
  let width = 9;
  let nextCode = LZW_FIRST_CODE;
  let previous: Uint8Array | null = null;
  let buffer = 0;
  let bits = 0;
  let cursor = 0;
  for (;;) {
    while (bits < width && cursor < encoded.length) {
      buffer = (buffer << 8) | encoded[cursor];
      cursor += 1;
      bits += 8;
    }
    // Fewer bits than one code left is the zero padding of the final byte.
    if (bits < width) break;
    const code = (buffer >>> (bits - width)) & ((1 << width) - 1);
    bits -= width;
    buffer &= (1 << bits) - 1;
    if (code === LZW_CLEAR_CODE) {
      table.length = 0;
      width = 9;
      nextCode = LZW_FIRST_CODE;
      previous = null;
      continue;
    }
    if (code === LZW_END_CODE) break;
    let entry: Uint8Array;
    if (code < LZW_CLEAR_CODE) {
      entry = Uint8Array.of(code);
    } else if (code - LZW_FIRST_CODE < table.length) {
      entry = table[code - LZW_FIRST_CODE];
    } else if (code === nextCode && previous !== null) {
      // The producer used an entry in the same code that defines it.
      entry = concatBytes([previous, previous.slice(0, 1)]);
    } else {
      fail("unsupported-document", "LZW stream holds an undefined code.");
    }
    for (const value of entry) decoded.push(value);
    if (decoded.length > MAX_DECODED_STREAM_BYTES) {
      fail("unsupported-document", "LZW stream is too large.");
    }
    if (previous !== null && nextCode < LZW_MAX_CODE) {
      table.push(concatBytes([previous, entry.slice(0, 1)]));
      nextCode += 1;
      if (width < 12 && ((nextCode + 1) & nextCode) === 0) width += 1;
    }
    previous = entry;
  }
  return Uint8Array.from(decoded);
}

function encodeLzwStream(decoded: Uint8Array): Uint8Array {
  const encoded: number[] = [];
  let buffer = 0;
  let bits = 0;
  let width = 9;
  const emit = (code: number): void => {
    buffer = (buffer << width) | code;
    bits += width;
    while (bits >= 8) {
      encoded.push((buffer >>> (bits - 8)) & 0xff);
      bits -= 8;
      buffer &= (1 << bits) - 1;
    }
  };
  let nextCode = LZW_FIRST_CODE;
  let table = new Map<number, number>();
  emit(LZW_CLEAR_CODE);
  if (decoded.length !== 0) {
    let sequence = decoded[0];
    for (let index = 1; index < decoded.length; index += 1) {
      const value = decoded[index];
      const extended = table.get(sequence * 256 + value);
      if (extended !== undefined) {
        sequence = extended;
        continue;
      }
      emit(sequence);
      table.set(sequence * 256 + value, nextCode);
      nextCode += 1;
      if (nextCode >= LZW_MAX_CODE - 1) {
        // One assignment before a 13th bit would be needed: clear at the
        // width both sides still agree on, then start the table over.
        emit(LZW_CLEAR_CODE);
        table = new Map<number, number>();
        nextCode = LZW_FIRST_CODE;
        width = 9;
      } else if (nextCode === 1 << width) {
        width += 1;
      }
      sequence = value;
    }
    emit(sequence);
  }
  emit(LZW_END_CODE);
  if (bits !== 0) encoded.push((buffer << (8 - bits)) & 0xff);
  return Uint8Array.from(encoded);
}

function decodeAscii85Stream(encoded: Uint8Array): Uint8Array {
  const decoded: number[] = [];
  let group: number[] = [];
  let cursor = 0;
  let terminated = false;

  const appendGroup = (digits: readonly number[], byteCount: number): void => {
    let value = 0;
    for (const digit of digits) value = value * 85 + digit;
    if (!Number.isSafeInteger(value) || value > 0xffff_ffff) {
      fail("unsupported-document", "ASCII85 stream group is out of range.");
    }
    for (let shift = 24; shift >= 32 - byteCount * 8; shift -= 8) {
      decoded.push((value >>> shift) & 0xff);
      if (decoded.length > MAX_DECODED_STREAM_BYTES) {
        fail("unsupported-document", "ASCII85 stream is too large.");
      }
    }
  };

  while (cursor < encoded.length) {
    const value = encoded[cursor++];
    if (isPdfWhitespaceByte(value)) continue;
    if (value === 0x7e) {
      if (cursor >= encoded.length || encoded[cursor++] !== 0x3e) {
        fail("unsupported-document", "ASCII85 stream terminator is malformed.");
      }
      terminated = true;
      break;
    }
    if (value === 0x7a) {
      if (group.length !== 0) {
        fail("unsupported-document", "ASCII85 zero shorthand splits a group.");
      }
      decoded.push(0, 0, 0, 0);
      if (decoded.length > MAX_DECODED_STREAM_BYTES) {
        fail("unsupported-document", "ASCII85 stream is too large.");
      }
      continue;
    }
    if (value < 0x21 || value > 0x75) {
      fail("unsupported-document", "ASCII85 stream contains an invalid byte.");
    }
    group.push(value - 0x21);
    if (group.length === 5) {
      appendGroup(group, 4);
      group = [];
    }
  }
  if (!terminated) {
    fail("unsupported-document", "ASCII85 stream terminator is missing.");
  }
  while (cursor < encoded.length) {
    if (!isPdfWhitespaceByte(encoded[cursor++])) {
      fail("unsupported-document", "ASCII85 stream has trailing data.");
    }
  }
  if (group.length === 1) {
    fail("unsupported-document", "ASCII85 final group is incomplete.");
  }
  if (group.length > 1) {
    const byteCount = group.length - 1;
    while (group.length < 5) group.push(84);
    appendGroup(group, byteCount);
  }
  return Uint8Array.from(decoded);
}

function encodeAscii85Stream(decoded: Uint8Array): Uint8Array {
  const encoded: number[] = [];
  for (let offset = 0; offset < decoded.length; offset += 4) {
    const count = Math.min(4, decoded.length - offset);
    const value =
      ((decoded[offset] ?? 0) * 0x100_0000 +
        (decoded[offset + 1] ?? 0) * 0x1_0000 +
        (decoded[offset + 2] ?? 0) * 0x100 +
        (decoded[offset + 3] ?? 0)) >>>
      0;
    if (count === 4 && value === 0) {
      encoded.push(0x7a);
      continue;
    }
    const digits = new Array<number>(5);
    let remainder = value;
    for (let index = 4; index >= 0; index -= 1) {
      digits[index] = (remainder % 85) + 0x21;
      remainder = Math.floor(remainder / 85);
    }
    encoded.push(...digits.slice(0, count + 1));
  }
  encoded.push(0x7e, 0x3e);
  return Uint8Array.from(encoded);
}








const INFLATED_STREAMS = 2;
const INFLATED_STREAM_MIN_BYTES = 64 * 1024;
const inflatedStreams: { encoded: Uint8Array; decoded: Uint8Array }[] = [];
function sameBytes(left: Uint8Array, right: Uint8Array): boolean {
  if (left.length !== right.length) return false;
  for (let index = 0; index < left.length; index += 1) if (left[index] !== right[index]) return false;
  return true;
}
function decodeFlateStream(encoded: Uint8Array): Uint8Array {
  if (encoded.length < INFLATED_STREAM_MIN_BYTES) return decodeFlateStreamUncached(encoded);
  const at = inflatedStreams.findIndex((entry) => sameBytes(entry.encoded, encoded));
  if (at >= 0) {
    const [entry] = inflatedStreams.splice(at, 1);
    inflatedStreams.unshift(entry);
    return Uint8Array.from(entry.decoded);
  }
  const decoded = decodeFlateStreamUncached(encoded);
  inflatedStreams.unshift({ encoded: Uint8Array.from(encoded), decoded: Uint8Array.from(decoded) });
  inflatedStreams.length = Math.min(inflatedStreams.length, INFLATED_STREAMS);
  return decoded;
}
function decodeFlateStreamUncached(encoded: Uint8Array): Uint8Array {
  if (
    encoded.length < 4 ||
    (encoded[0] & 0x0f) !== 8 ||
    encoded[0] >>> 4 > 7 ||
    ((encoded[0] << 8) | encoded[1]) % 31 !== 0 ||
    (encoded[1] & 0x20) !== 0
  ) {
    fail("unsupported-document", "Flate stream header is malformed.");
  }
  const chunks: Uint8Array[] = [];
  let decodedLength = 0;
  let complete = false;
  try {
    const inflater = new Unzlib((chunk, final) => {
      decodedLength += chunk.length;
      if (decodedLength > MAX_DECODED_STREAM_BYTES) {
        throw new Error("decoded stream exceeds limit");
      }
      if (chunk.length) chunks.push(Uint8Array.from(chunk));
      if (final) complete = true;
    });
    inflater.push(Uint8Array.from(encoded), true);
  } catch {
    // An empty stream written without its trailing checksum (a header and
    // an empty final block, 4 bytes: Word/Distiller pages and forms that
    
    // and WPS read it as empty. Mirror of the server's rule.
    if (encoded.length <= 8) {
      let raw: Uint8Array | null = null;
      try {
        raw = inflateSync(encoded.subarray(2));
      } catch {
        raw = null;
      }
      if (raw && raw.length === 0) return new Uint8Array(0);
    }
    const whole = decodeFlateStreamWithShortChecksum(encoded);
    if (whole) return whole;
    fail("unsupported-document", "Flate stream is malformed or too large.");
  }
  if (!complete) {
    fail("unsupported-document", "Flate stream is incomplete.");
  }
  const decoded = concatBytes(chunks);
  // Many producers count the end-of-line that separates the stream data from
  // `endstream` in /Length, so the declared span runs a byte or two past the
  // zlib stream and its trailing checksum is no longer the last four bytes.
  // The checksum still has to match the decoded bytes exactly -- only where it
  // sits is in question, and only across that end-of-line: every byte after it
  // must be CR or LF. Replay applies the same rule.
  const actual = adler32(decoded);
  let checksumOffset = -1;
  for (let trailing = 0; trailing <= MAX_FLATE_TRAILING_BYTES; trailing += 1) {
    const candidate = encoded.length - 4 - trailing;
    if (candidate < 0) break;
    const value =
      ((encoded[candidate] << 24) |
        (encoded[candidate + 1] << 16) |
        (encoded[candidate + 2] << 8) |
        encoded[candidate + 3]) >>>
      0;
    if (
      value === actual &&
      encoded.subarray(candidate + 4).every((byte) => byte === 0x0d || byte === 0x0a)
    ) {
      checksumOffset = candidate;
      break;
    }
  }
  if (checksumOffset < 0) {
    fail("unsupported-document", "Flate stream checksum is invalid.");
  }
  return decoded;
}











function decodeFlateStreamWithShortChecksum(encoded: Uint8Array): Uint8Array | null {
  const inflated = (end: number): Uint8Array | null => {
    if (end <= 2) return null;
    const chunks: Uint8Array[] = [];
    let length = 0;
    try {
      const inflater = new Inflate((chunk) => {
        length += chunk.length;
        if (length > MAX_DECODED_STREAM_BYTES) throw new Error("decoded stream exceeds limit");
        if (chunk.length) chunks.push(Uint8Array.from(chunk));
      });
      inflater.push(Uint8Array.from(encoded.subarray(2, end)), true);
    } catch {
      return null;
    }
    return concatBytes(chunks);
  };
  for (let present = 3; present >= 1; present -= 1) {
    const end = encoded.length - present;
    const decoded = inflated(end);
    if (!decoded) continue;
    // The data's last byte is the one before the checksum: with a byte
    // fewer it is no whole stream.
    if (inflated(end - 1)) return null;
    const checksum = adler32(decoded);
    const expected = [checksum >>> 24, (checksum >>> 16) & 0xff, (checksum >>> 8) & 0xff];
    return encoded.subarray(end).every((byte, index) => byte === expected[index]) ? decoded : null;
  }
  return null;
}

function adler32(bytes: Uint8Array): number {
  let first = 1;
  let second = 0;
  for (let offset = 0; offset < bytes.length; offset += 5_552) {
    const end = Math.min(bytes.length, offset + 5_552);
    for (let index = offset; index < end; index += 1) {
      first += bytes[index];
      second += first;
    }
    first %= 65_521;
    second %= 65_521;
  }
  return ((second << 16) | first) >>> 0;
}

function encodeDeterministicFlateStream(decoded: Uint8Array): Uint8Array {
  if (decoded.length > MAX_DECODED_STREAM_BYTES) {
    fail("unsupported-document", "Decoded PDF stream is too large.");
  }
  const blockCount = Math.max(1, Math.ceil(decoded.length / 65_535));
  const encoded = new Uint8Array(2 + blockCount * 5 + decoded.length + 4);
  encoded[0] = 0x78;
  encoded[1] = 0x01;
  let sourceOffset = 0;
  let targetOffset = 2;
  for (let block = 0; block < blockCount; block += 1) {
    const length = Math.min(65_535, decoded.length - sourceOffset);
    const complement = 0xffff ^ length;
    encoded[targetOffset] = block === blockCount - 1 ? 0x01 : 0x00;
    encoded[targetOffset + 1] = length & 0xff;
    encoded[targetOffset + 2] = length >>> 8;
    encoded[targetOffset + 3] = complement & 0xff;
    encoded[targetOffset + 4] = complement >>> 8;
    targetOffset += 5;
    encoded.set(
      decoded.subarray(sourceOffset, sourceOffset + length),
      targetOffset,
    );
    sourceOffset += length;
    targetOffset += length;
  }
  const checksum = adler32(decoded);
  encoded[targetOffset] = checksum >>> 24;
  encoded[targetOffset + 1] = (checksum >>> 16) & 0xff;
  encoded[targetOffset + 2] = (checksum >>> 8) & 0xff;
  encoded[targetOffset + 3] = checksum & 0xff;
  return encoded;
}

function ascii(value: string): Uint8Array {
  if ([...value].some((character) => character.charCodeAt(0) > 0x7f)) {
    fail("invalid-request", "Native PDF syntax must be ASCII.");
  }
  return ASCII.encode(value);
}

function asLatin1(bytes: Uint8Array): string {
  return LATIN1.decode(bytes);
}

/** One source decoded once, kept while the same bytes are read again. An
 * object read that is not handed a decoded source decodes the whole file to
 * find the object's offset in it, and a plan reads objects of the same source
 * hundreds of times over: on a 800 KB page the text planner spent about a
 * second of every text-box opening re-decoding those same bytes. Sources are
 * never written in place, and each entry dies with its buffer. Small buffers
 * (a virtual object read) are decoded directly, so this holds sources only. */
const decodedSources = new WeakMap<Uint8Array, Fingerprinted<string>>();

/** What is kept against a caller's buffer, and the bytes it was kept for.
 *
 * A buffer is the key, but a caller may write into its own buffer and read
 * it again (the batch reference counter's contract), so what was kept is only
 * as good as the bytes it was made from. Every read of a source begins by
 * parsing it (`parsePdf`), and that is where the buffer is summed and
 * anything kept for other bytes is dropped; the reads inside one parse then
 * find their entries by identity alone. Adler-32 over the whole buffer is a
 * millisecond a megabyte, once per read, against the decode, parse and
 * inflate it stands in for; it is a staleness check, never an identity -- the
 * hashes that bind a record to its source are still taken over the bytes. */
interface Fingerprinted<T> {
  readonly fingerprint: number;
  readonly value: T;
}

const FINGERPRINTED_SOURCE_MIN_BYTES = 4096;

function bufferFingerprint(bytes: Uint8Array): number {
  // A staleness check, never an identity: 32-bit mixing over whole words
  // where the view allows it (a byte loop over a 5MB source took 25ms, four
  // times a save), the odd bytes at the end folded in one at a time.
  let hash = 0x811c9dc5 ^ bytes.length;
  let index = 0;
  if (bytes.byteOffset % 4 === 0) {
    const words = new Uint32Array(bytes.buffer, bytes.byteOffset, bytes.length >>> 2);
    // An indexed loop: iterating the view (for...of) ran five times slower
    // on a 6MB source, and a save fingerprints its source and its candidate.
    const count = words.length;
    for (let word = 0; word < count; word += 1) {
      hash = Math.imul(hash ^ words[word], 0x01000193);
      hash ^= hash >>> 15;
    }
    index = count << 2;
  }
  for (; index < bytes.length; index += 1) {
    hash = Math.imul(hash ^ bytes[index], 0x01000193);
  }
  return hash >>> 0;
}

/** Drops what was kept for `bytes` under other contents, and says what the
 * buffer sums to now. */
function freshSourceFingerprint(bytes: Uint8Array): number {
  const fingerprint = bufferFingerprint(bytes);
  for (const kept of [decodedSources, parsedSources] as const) {
    const entry = kept.get(bytes);
    if (entry !== undefined && entry.fingerprint !== fingerprint) kept.delete(bytes);
  }
  return fingerprint;
}

/** Candidates proven to begin with a source's exact bytes, mapped to that
 * source: an append-only save's candidate is its source plus one increment,
 * and a byte decodes the same wherever it sits, so the candidate's text is the
 * source's own followed by the increment's. */
const candidateSourcePrefixes = new WeakMap<Uint8Array, Uint8Array>();

function decodedSourceLatin1(bytes: Uint8Array): string {
  if (bytes.length < FINGERPRINTED_SOURCE_MIN_BYTES) return asLatin1(bytes);
  const cached = decodedSources.get(bytes);
  if (cached !== undefined) return cached.value;
  const prefix = candidateSourcePrefixes.get(bytes);
  const source = prefix && prefix.length <= bytes.length && bytesEqual(bytes.subarray(0, prefix.length), prefix)
    ? decodedSourceLatin1(prefix) + asLatin1(bytes.subarray(prefix.length))
    : asLatin1(bytes);
  decodedSources.set(bytes, { fingerprint: bufferFingerprint(bytes), value: source });
  return source;
}

function uint32BigEndian(bytes: Uint8Array, offset: number): number {
  if (offset < 0 || offset + 4 > bytes.length) {
    fail("invalid-request", "Replacement image integer is out of bounds.");
  }
  return (
    bytes[offset] * 0x1000000 +
    (bytes[offset + 1] << 16) +
    (bytes[offset + 2] << 8) +
    bytes[offset + 3]
  );
}

function crc32(bytes: Uint8Array): number {
  let value = 0xffffffff;
  for (const byte of bytes) {
    value ^= byte;
    for (let bit = 0; bit < 8; bit += 1) {
      value = (value >>> 1) ^ (0xedb88320 & -(value & 1));
    }
  }
  return (value ^ 0xffffffff) >>> 0;
}

function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000));
  }
  return btoa(binary);
}

function base64ToBytes(value: string): Uint8Array {
  if (
    value.length < 4 ||
    value.length > Math.ceil(MAX_MRAI_NATIVE_EMBEDDED_IMAGE_BYTES / 3) * 4 ||
    !/^[A-Za-z0-9+/]+={0,2}$/.test(value) ||
    value.length % 4 !== 0
  ) {
    fail("invalid-request", "Replacement image payload is invalid.");
  }
  let decoded: string;
  try {
    decoded = atob(value);
  } catch {
    fail("invalid-request", "Replacement image payload is invalid.");
  }
  const bytes = Uint8Array.from(decoded, (character) =>
    character.charCodeAt(0),
  );
  if (
    bytes.length < 1 ||
    bytes.length > MAX_MRAI_NATIVE_EMBEDDED_IMAGE_BYTES ||
    bytesToBase64(bytes) !== value
  ) {
    fail("invalid-request", "Replacement image payload is invalid.");
  }
  return bytes;
}

function exactReplacementImageDimension(value: number, label: string): number {
  if (!Number.isSafeInteger(value) || value < 1 || value > 100_000) {
    fail("invalid-request", `Replacement image ${label} is unsupported.`);
  }
  return value;
}

async function replacementImageDataFromJpeg(
  bytes: Uint8Array,
): Promise<PdfNativeEmbeddedImageReplacementData> {
  if (
    bytes.length < 12 ||
    bytes[0] !== 0xff ||
    bytes[1] !== 0xd8 ||
    bytes.at(-2) !== 0xff ||
    bytes.at(-1) !== 0xd9
  ) {
    fail("invalid-request", "Replacement JPEG is malformed.");
  }
  let cursor = 2;
  let frame: { width: number; height: number; components: 1 | 3 } | null = null;
  while (cursor < bytes.length - 2) {
    if (bytes[cursor] !== 0xff) {
      fail("invalid-request", "Replacement JPEG marker is malformed.");
    }
    while (cursor < bytes.length && bytes[cursor] === 0xff) cursor += 1;
    const marker = bytes[cursor++];
    if (marker === 0xd9) break;
    if (marker === 0xda) {
      if (!frame) fail("invalid-request", "Replacement JPEG frame is missing.");
      break;
    }
    if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) continue;
    if (cursor + 2 > bytes.length) {
      fail("invalid-request", "Replacement JPEG segment is truncated.");
    }
    const length = (bytes[cursor] << 8) | bytes[cursor + 1];
    if (length < 2 || cursor + length > bytes.length) {
      fail("invalid-request", "Replacement JPEG segment is malformed.");
    }
    const dataStart = cursor + 2;
    const dataEnd = cursor + length;
    if (
      marker === 0xe2 &&
      asLatin1(bytes.subarray(dataStart, Math.min(dataEnd, dataStart + 12))) ===
        "ICC_PROFILE\u0000"
    ) {
      fail("unsupported-document", "ICC replacement images are preserve-only.");
    }
    if (marker === 0xc0 || marker === 0xc2) {
      if (frame || length < 8 || bytes[dataStart] !== 8) {
        fail("unsupported-document", "Replacement JPEG frame is unsupported.");
      }
      const height = exactReplacementImageDimension(
        (bytes[dataStart + 1] << 8) | bytes[dataStart + 2],
        "height",
      );
      const width = exactReplacementImageDimension(
        (bytes[dataStart + 3] << 8) | bytes[dataStart + 4],
        "width",
      );
      const components = bytes[dataStart + 5];
      if (components !== 1 && components !== 3) {
        fail("unsupported-document", "CMYK JPEG replacement is unsupported.");
      }
      frame = { width, height, components };
    }
    cursor = dataEnd;
  }
  if (
    !frame ||
    frame.width * frame.height > MAX_MRAI_NATIVE_EMBEDDED_IMAGE_PIXELS
  ) {
    fail(
      "unsupported-document",
      "Replacement JPEG dimensions are unsupported.",
    );
  }
  return Object.freeze({
    mimeType: "image/jpeg",
    width: frame.width,
    height: frame.height,
    colorSpace: frame.components === 1 ? "DeviceGray" : "DeviceRGB",
    bitsPerComponent: 8,
    filter: "DCTDecode",
    decodeParms: null,
    encodedDataBase64: bytesToBase64(bytes),
    encodedSha256: await sha256PdfBytes(bytes),
  });
}

async function replacementImageDataFromPng(
  bytes: Uint8Array,
): Promise<PdfNativeEmbeddedImageReplacementData> {
  const signature = [137, 80, 78, 71, 13, 10, 26, 10];
  if (
    bytes.length < 45 ||
    signature.some((byte, index) => bytes[index] !== byte)
  ) {
    fail("invalid-request", "Replacement PNG is malformed.");
  }
  let cursor = 8;
  let width = 0;
  let height = 0;
  let colors: 1 | 3 | null = null;
  let hasAlpha = false;
  let sawHeader = false;
  let sawData = false;
  let dataEnded = false;
  let sawEnd = false;
  const imageData: Uint8Array[] = [];
  while (cursor < bytes.length) {
    const length = uint32BigEndian(bytes, cursor);
    const typeStart = cursor + 4;
    const dataStart = typeStart + 4;
    const dataEnd = dataStart + length;
    const crcEnd = dataEnd + 4;
    if (
      length > MAX_MRAI_NATIVE_EMBEDDED_IMAGE_BYTES ||
      crcEnd > bytes.length
    ) {
      fail("invalid-request", "Replacement PNG chunk is malformed.");
    }
    const typeBytes = bytes.subarray(typeStart, dataStart);
    const type = asLatin1(typeBytes);
    const actualCrc = uint32BigEndian(bytes, dataEnd);
    if (crc32(bytes.subarray(typeStart, dataEnd)) !== actualCrc) {
      fail("invalid-request", "Replacement PNG checksum is invalid.");
    }
    if (!/^[A-Za-z]{4}$/.test(type)) {
      fail("invalid-request", "Replacement PNG chunk type is malformed.");
    }
    const critical = typeBytes[0] >= 65 && typeBytes[0] <= 90;
    if (type === "IHDR") {
      if (sawHeader || cursor !== 8 || length !== 13) {
        fail("invalid-request", "Replacement PNG header is ambiguous.");
      }
      width = exactReplacementImageDimension(
        uint32BigEndian(bytes, dataStart),
        "width",
      );
      height = exactReplacementImageDimension(
        uint32BigEndian(bytes, dataStart + 4),
        "height",
      );
      const bitDepth = bytes[dataStart + 8];
      const colorType = bytes[dataStart + 9];
      if (
        bitDepth !== 8 ||
        ![0, 2, 4, 6].includes(colorType) ||
        bytes[dataStart + 10] !== 0 ||
        bytes[dataStart + 11] !== 0 ||
        bytes[dataStart + 12] !== 0
      ) {
        fail("unsupported-document", "Replacement PNG format is unsupported.");
      }
      colors = colorType === 0 || colorType === 4 ? 1 : 3;
      hasAlpha = colorType === 4 || colorType === 6;
      sawHeader = true;
    } else if (type === "IDAT") {
      if (!sawHeader || dataEnded || length < 1) {
        fail("invalid-request", "Replacement PNG image data is malformed.");
      }
      sawData = true;
      imageData.push(Uint8Array.from(bytes.subarray(dataStart, dataEnd)));
    } else if (type === "IEND") {
      if (
        !sawHeader ||
        !sawData ||
        sawEnd ||
        length !== 0 ||
        crcEnd !== bytes.length
      ) {
        fail("invalid-request", "Replacement PNG terminator is malformed.");
      }
      sawEnd = true;
    } else {
      if (sawData) dataEnded = true;
      if (
        critical ||
        type === "iCCP" ||
        type === "tRNS" ||
        type === "eXIf" ||
        ["acTL", "fcTL", "fdAT"].includes(type)
      ) {
        fail("unsupported-document", "Replacement PNG profile is unsupported.");
      }
    }
    cursor = crcEnd;
  }
  if (
    !sawEnd ||
    !colors ||
    width * height > MAX_MRAI_NATIVE_EMBEDDED_IMAGE_PIXELS
  ) {
    fail("unsupported-document", "Replacement PNG dimensions are unsupported.");
  }
  const encoded = concatBytes(imageData);
  if (
    encoded.length < 1 ||
    encoded.length > MAX_MRAI_NATIVE_EMBEDDED_IMAGE_BYTES
  ) {
    fail("unsupported-document", "Replacement PNG payload is too large.");
  }
  const decoded = decodeFlateStream(encoded);
  const rowLength = width * (colors + (hasAlpha ? 1 : 0)) + 1;
  if (
    !Number.isSafeInteger(rowLength) ||
    decoded.length !== rowLength * height
  ) {
    fail("invalid-request", "Replacement PNG scanlines are malformed.");
  }
  for (let row = 0; row < height; row += 1) {
    if (decoded[row * rowLength] > 4) {
      fail("invalid-request", "Replacement PNG row filter is malformed.");
    }
  }
  if (hasAlpha) {
    if (width * height * 4 > MAX_DECODED_STREAM_BYTES) {
      fail(
        "invalid-request",
        "PNG color and alpha planes exceed the native byte budget.",
      );
    }
    // Reuse the PNG decoder already used for images throughout the platform.
    // This temporary context never receives or rewrites the source PDF.
    const {
      PngEmbedder,
      PDFContext,
      PDFName,
      PDFRawStream,
      decodePDFRawStream,
    } = await import("pdf-lib");
    const context = PDFContext.create();
    const embedder = await PngEmbedder.for(bytes);
    const image = context.lookup(await embedder.embedIntoContext(context));
    if (!(image instanceof PDFRawStream)) {
      fail("invalid-request", "PNG color stream could not be decoded.");
    }
    const rgb = decodePDFRawStream(image).decode();
    const maskRef = image.dict.get(PDFName.of("SMask"));
    const maskStream = maskRef ? context.lookup(maskRef) : null;
    if (maskRef && !(maskStream instanceof PDFRawStream)) {
      fail("invalid-request", "PNG alpha stream could not be decoded.");
    }
    const alpha =
      maskStream instanceof PDFRawStream
        ? decodePDFRawStream(maskStream).decode()
        : null;
    if (
      embedder.width !== width ||
      embedder.height !== height ||
      rgb.length !== width * height * 3 ||
      (alpha && alpha.length !== width * height)
    ) {
      fail("invalid-request", "PNG channel dimensions are inconsistent.");
    }
    const plane = async (pixels: Uint8Array, components: number) => {
      const stride = width * components;
      if ((stride + 1) * height > MAX_DECODED_STREAM_BYTES) {
        fail("invalid-request", "PNG channel exceeds the native byte budget.");
      }
      const rows = new Uint8Array((stride + 1) * height);
      for (let row = 0; row < height; row += 1) {
        rows.set(
          pixels.subarray(row * stride, (row + 1) * stride),
          row * (stride + 1) + 1,
        );
      }
      const encodedPlane = zlibSync(rows);
      return Object.freeze({
        encodedDataBase64: bytesToBase64(encodedPlane),
        encodedSha256: await sha256PdfBytes(encodedPlane),
      });
    };
    const colorPlane = await plane(rgb, 3);
    const softMask = alpha ? await plane(alpha, 1) : undefined;
    if (
      colorPlane.encodedDataBase64.length +
        (softMask?.encodedDataBase64.length ?? 0) >
      Math.ceil(MAX_MRAI_NATIVE_EMBEDDED_IMAGE_BYTES / 3) * 4
    ) {
      fail("invalid-request", "PNG channels exceed the native payload budget.");
    }
    return Object.freeze({
      mimeType: "image/png",
      width,
      height,
      colorSpace: "DeviceRGB",
      bitsPerComponent: 8,
      filter: "FlateDecode",
      decodeParms: Object.freeze({
        predictor: 15,
        colors: 3,
        bitsPerComponent: 8,
        columns: width,
      }),
      ...colorPlane,
      ...(softMask ? { softMask } : {}),
    });
  }
  return Object.freeze({
    mimeType: "image/png",
    width,
    height,
    colorSpace: colors === 1 ? "DeviceGray" : "DeviceRGB",
    bitsPerComponent: 8,
    filter: "FlateDecode",
    decodeParms: Object.freeze({
      predictor: 15,
      colors,
      bitsPerComponent: 8,
      columns: width,
    }),
    encodedDataBase64: bytesToBase64(encoded),
    encodedSha256: await sha256PdfBytes(encoded),
  });
}

export async function inspectMraiNativeEmbeddedImageReplacement(
  bytes: Uint8Array,
  mimeType: string,
): Promise<PdfNativeEmbeddedImageReplacementData> {
  if (
    !isUint8Array(bytes) ||
    bytes.length < 1 ||
    bytes.length > MAX_MRAI_NATIVE_EMBEDDED_IMAGE_BYTES
  ) {
    fail("invalid-request", "Replacement image exceeds the native limit.");
  }
  return mimeType === "image/jpeg"
    ? replacementImageDataFromJpeg(bytes)
    : mimeType === "image/png"
      ? replacementImageDataFromPng(bytes)
      : fail("unsupported-document", "Replacement image type is unsupported.");
}

function exactPositiveInteger(
  value: number,
  label: string,
  allowZero = false,
): number {
  if (
    !Number.isSafeInteger(value) ||
    value < (allowZero ? 0 : 1) ||
    value > 9_999_999_999
  ) {
    fail("invalid-request", `Invalid ${label}.`);
  }
  return value;
}

function exactGeneration(value: number): number {
  if (!Number.isSafeInteger(value) || value < 0 || value > 65_535) {
    fail("invalid-request", "Invalid PDF object generation.");
  }
  return value;
}

function exactSha256(value: string, label: string): string {
  const normalized = value.toLowerCase();
  if (!SHA256_PATTERN.test(normalized)) {
    fail("invalid-request", `Invalid ${label} SHA-256.`);
  }
  return normalized;
}

/** The bytes a plan reads from: the caller's own array when it is a whole,
 * unshared buffer, else a private copy. Every plan of one edit (the native
 * route, the font-reuse check, the writer's proof) then reads the same
 * array, and the per-source caches (xref, digest, page graphs) hit; a copy
 * per plan made a 1200-page book walk its pages once per plan. Sources
 * are never written in place. */
export function ownedPdfSourceBytes(bytes: Uint8Array): Uint8Array {
  const shared =
    typeof SharedArrayBuffer !== "undefined" && bytes.buffer instanceof SharedArrayBuffer;
  return !shared && bytes.byteOffset === 0 && bytes.byteLength === bytes.buffer.byteLength
    ? bytes
    : Uint8Array.from(bytes);
}


function assertSourceEnvelope(bytes: Uint8Array): string {
  if (bytes.length === 0 || bytes.length > MAX_MRAI_NATIVE_SOURCE_BYTES) {
    fail("unsupported-document", "PDF source size is unsupported.");
  }
  const source = decodedSourceLatin1(bytes);
  if (!source.startsWith("%PDF-1.") && !source.startsWith("%PDF-2.")) {
    fail("unsupported-document", "PDF header is missing or unsupported.");
  }

  // Encryption is not decided here: a file is encrypted by a trailer on its
  // chain, not by the word (refuseEncryptionNamedAnywhere).
  const blocked: readonly [RegExp, string][] = [
    // A certification signature carries a rule about what may still change,
    // and text is never part of it. An ordinary approval signature carries
    // no such rule: it signs the revision it was made over, and a native
    // edit only appends a new one, so those bytes and that signature stand
    // exactly as they were.
    // /Perms on its own carries usage rights (UR3), a grant a reader honours
    // and an edit merely invalidates; only a DocMDP transform states what
    // may change, and a /Perms that certifies names it.
    [
      /\/DocMDP\b/,
      "Certified PDFs are preserve-only.",
    ],
    // /XFA is decided after the xref is parsed, not here: whether the reader
    // draws this file's page content or regenerates it from the XFA template
    // is a fact of the catalog, and the key itself is usually compressed into
    // an object stream anyway. See assertXfaRenderingAuthority.
  ];
  for (const [pattern, message] of blocked) {
    if (pattern.test(source)) fail("unsupported-document", message);
  }
  return source;
}

/**
 * A file is encrypted when the trailer or the cross-reference stream
 * dictionary of a section on its chain -- the newest and every one its /Prev
 * leads to -- names /Encrypt; each section refuses its own (parseClassicXref,
 * parseXrefStream). The word elsewhere in the file encrypts nothing: a file
 * decrypted by a tool that wrote a new table may still hold the superseded
 * one that names the key (walk 0454), and a page may print the word.
 *
 * That holds only for the chain the file records. Where it is not the one
 * read -- a section had to be recovered, or the file is a complete PDF
 * followed by other bytes -- readers rebuild the file each their own way and
 * one may come upon a section this parser passed over, so there the word
 * anywhere still refuses (mirror of replay's `_refuse_encryption_named_anywhere`).
 */
function refuseEncryptionNamedAnywhere(source: string): void {
  if (/\/Encrypt\b/.test(source)) {
    fail("unsupported-document", "Encrypted PDFs are preserve-only.");
  }
}

function isPdfWhitespace(character: string | undefined): boolean {
  if (!character) return false;
  const code = character.charCodeAt(0);
  return (
    code === 0x00 ||
    code === 0x09 ||
    code === 0x0a ||
    code === 0x0c ||
    code === 0x0d ||
    code === 0x20
  );
}

function skipWhitespace(source: string, start: number): number {
  let cursor = start;
  while (cursor < source.length) {
    const character = source[cursor];
    if (isPdfWhitespace(character)) {
      cursor += 1;
      continue;
    }
    if (character === "%") {
      while (
        cursor < source.length &&
        source[cursor] !== "\n" &&
        source[cursor] !== "\r"
      ) {
        cursor += 1;
      }
      continue;
    }
    break;
  }
  return cursor;
}

function readLine(
  source: string,
  start: number,
): { readonly line: string; readonly next: number } {
  let end = start;
  while (end < source.length && source[end] !== "\n" && source[end] !== "\r") {
    end += 1;
  }
  let next = end;
  if (source[next] === "\r") next += 1;
  if (source[next] === "\n") next += 1;
  return { line: source.slice(start, end).trim(), next };
}

function dictionaryEnd(source: string, start: number): number {
  if (source.slice(start, start + 2) !== "<<") {
    fail("unsupported-document", "PDF dictionary is malformed.");
  }
  let depth = 0;
  let literalDepth = 0;
  let escaped = false;
  let inHex = false;
  for (let cursor = start; cursor < source.length; cursor += 1) {
    const character = source[cursor];
    const next = source[cursor + 1];
    if (literalDepth > 0) {
      if (escaped) {
        escaped = false;
      } else if (character === "\\") {
        escaped = true;
      } else if (character === "(") {
        literalDepth += 1;
      } else if (character === ")") {
        literalDepth -= 1;
      }
      continue;
    }
    if (inHex) {
      if (character === ">") inHex = false;
      continue;
    }
    if (character === "%") {
      while (
        cursor < source.length &&
        source[cursor] !== "\n" &&
        source[cursor] !== "\r"
      ) {
        cursor += 1;
      }
      continue;
    }
    if (character === "(") {
      literalDepth = 1;
      continue;
    }
    if (character === "<" && next !== "<") {
      inHex = true;
      continue;
    }
    if (character === "<" && next === "<") {
      depth += 1;
      cursor += 1;
      continue;
    }
    if (character === ">" && next === ">") {
      depth -= 1;
      cursor += 1;
      if (depth === 0) return cursor + 1;
      if (depth < 0) break;
    }
  }
  fail("unsupported-document", "PDF dictionary is unterminated.");
}

function oneInteger(
  value: string,
  pattern: RegExp,
  label: string,
): number | null {
  const matches = [...value.matchAll(pattern)];
  if (matches.length === 0) return null;
  if (matches.length !== 1) {
    fail("unsupported-document", `PDF ${label} is ambiguous.`);
  }
  return Number(matches[0][1]);
}

/** Where the file says its newest cross-reference section is, or null when
 * it does not end as a PDF does. NUL is PDF whitespace; a writer may pad the
 * file after `%%EOF` with it. */
function finalStartXref(source: string): number | null {
  const match = /startxref\s+(\d+)\s+%%EOF[\s\x00]*$/.exec(source);
  return match ? Number(match[1]) : null;
}

function findTokenAfter(source: string, token: string, start: number): number {
  const index = source.indexOf(token, start);
  if (index < 0) {
    fail("unsupported-document", `PDF ${token} boundary is missing.`);
  }
  return index;
}

/** Where a stream really ends when its /Length does not say: at the one
 * `endstream` the object's own `endobj` follows. Null when the object holds
 * no such terminator, which leaves the stream unreadable. */
function soleStreamTerminator(
  source: string,
  streamStart: number,
): Readonly<{ streamEnd: number; terminator: number }> | null {
  let cursor = streamStart;
  while (true) {
    const found = source.indexOf("endstream", cursor);
    if (found < 0) return null;
    const endobj = skipWhitespace(source, found + 9);
    if (source.slice(endobj, endobj + 6) === "endobj") {
      let end = found;
      if (source.slice(end - 2, end) === "\r\n") end -= 2;
      else if (source[end - 1] === "\n" || source[end - 1] === "\r") end -= 1;
      return end >= streamStart
        ? Object.freeze({ streamEnd: end, terminator: found })
        : null;
    }
    cursor = found + 9;
  }
}

function parseDirectObjectAtOffset(
  bytes: Uint8Array,
  offset: number,
  expected?: Readonly<{ objectNumber: number; generation: number }>,
  decodedSource?: string,
  resolveLength?: (
    reference: Readonly<{ objectNumber: number; generation: number }>,
  ) => number,
  /** How a stream is read: decoded; decoded where this engine can and
   * carried opaquely where it cannot; or raw -- a reader that only hashes
   * the bytes the producer wrote never inflates them. */
  streams: "decoded" | "opaque" | "raw" = "decoded",
): ParsedIndirectObject {
  const source = decodedSource ?? decodedSourceLatin1(bytes);
  if (!Number.isSafeInteger(offset) || offset < 0 || offset >= bytes.length) {
    fail("unsupported-document", "PDF xref object offset is out of range.");
  }
  const header = /^(\d+)\s+(\d+)\s+obj\b/.exec(source.slice(offset));
  if (!header) {
    fail("unsupported-document", "PDF xref does not own an indirect object.");
  }
  const objectNumber = Number(header[1]);
  const generation = Number(header[2]);
  if (
    !Number.isSafeInteger(objectNumber) ||
    objectNumber < 1 ||
    !Number.isSafeInteger(generation) ||
    generation < 0 ||
    generation > 65_535 ||
    (expected !== undefined &&
      (objectNumber !== expected.objectNumber ||
        generation !== expected.generation))
  ) {
    fail("unsupported-document", "PDF indirect object identity is invalid.");
  }

  const bodyStart = skipWhitespace(source, offset + header[0].length);
  let objectEnd: number;
  let dictionaryStart: number | null = null;
  let dictionaryStop: number | null = null;
  let lengthNumberStart: number | null = null;
  let lengthNumberEnd: number | null = null;
  let lengthKeyMissing = false;
  let lengthReference: Readonly<{
    objectNumber: number;
    generation: number;
  }> | null = null;
  let streamStart: number | null = null;
  let streamEnd: number | null = null;
  let streamEncoding:
    | "identity"
    | "flate"
    | "flate-predictor15"
    | "flate-png-predictor"
    | "ascii85-flate"
    | "ascii-hex"
    | "lzw"
    | "dct"
    | "ascii85-dct"
    | "opaque"
    | null = null;
  let encodedStreamBytes: Uint8Array | null = null;
  let streamBytes: Uint8Array | null = null;

  if (source.slice(bodyStart, bodyStart + 2) === "<<") {
    dictionaryStart = bodyStart;
    dictionaryStop = dictionaryEnd(source, bodyStart);
    const dictionary = source.slice(dictionaryStart, dictionaryStop);
    let afterDictionary = skipWhitespace(source, dictionaryStop);
    if (source.slice(afterDictionary, afterDictionary + 6) === "stream") {
      streamEncoding = parseStreamEncoding(dictionary, streams !== "decoded");
      const lengths = [
        ...dictionary.matchAll(/\/Length\s+(\d+)(?:\s+(\d+)\s+R)?/g),
      ];
      // A dictionary that names no /Length at all ends at its one
      // `endstream`, which is where readers end it (mirror of replay).
      lengthKeyMissing = lengths.length === 0 && !/\/Length\b/.test(dictionary);
      if (lengths.length !== 1 && !lengthKeyMissing) {
        fail(
          "unsupported-document",
          "PDF stream Length must be one direct integer.",
        );
      }
      // The span a revision replaces with its own length: the digits when the
      // producer wrote them inline, and the whole `N G R` when it stored them
      // in their own object. Writing the number inline there keeps the change
      // inside the object being rewritten anyway, and leaves the object that
      // held it untouched.
      let length: number;
      if (lengthKeyMissing) {
        // Nothing to patch: a revision writes the whole `/Length N ` here.
        lengthNumberStart = dictionaryStart + 2;
        lengthNumberEnd = dictionaryStart + 2;
        length = -1;
      } else {
      const valuePrefix = /^\/Length\s+/.exec(lengths[0][0])![0].length;
      lengthNumberStart = dictionaryStart + lengths[0].index! + valuePrefix;
      lengthNumberEnd =
        dictionaryStart + lengths[0].index! + lengths[0][0].length;
      if (lengths[0][2] !== undefined) {
        if (!resolveLength) {
          fail(
            "unsupported-document",
            "PDF stream Length must be one direct integer.",
          );
        }
        const reference = Object.freeze({
          objectNumber: Number(lengths[0][1]),
          generation: Number(lengths[0][2]),
        });
        assertObjectReference(reference.objectNumber, reference.generation);
        length = resolveLength(reference);
        lengthReference = reference;
      } else {
        length = Number(lengths[0][1]);
      }
      if (!Number.isSafeInteger(length) || length < 0) {
        fail("unsupported-document", "PDF stream Length is invalid.");
      }
      }

      afterDictionary += 6;
      if (source.slice(afterDictionary, afterDictionary + 2) === "\r\n") {
        streamStart = afterDictionary + 2;
      } else if (
        source[afterDictionary] === "\n" ||
        source[afterDictionary] === "\r"
      ) {
        streamStart = afterDictionary + 1;
      } else {
        fail("unsupported-document", "PDF stream EOL is malformed.");
      }
      streamEnd = lengthKeyMissing ? bytes.length + 1 : streamStart + length;
      let endstream = streamEnd;
      if (streamEnd <= bytes.length) {
        if (source.slice(endstream, endstream + 2) === "\r\n") endstream += 2;
        else if (source[endstream] === "\n" || source[endstream] === "\r") {
          endstream += 1;
        }
      }
      if (
        streamEnd > bytes.length ||
        source.slice(endstream, endstream + 9) !== "endstream"
      ) {
        // A /Length that misses its own `endstream` says nothing about where
        // the stream ends; the object's own terminator does, and that is
        // what readers take (mirrored by replay).
        const repaired = soleStreamTerminator(source, streamStart);
        if (repaired === null) {
          fail(
            "unsupported-document",
            streamEnd > bytes.length
              ? "PDF stream exceeds the source bytes."
              : "PDF stream terminator is malformed.",
          );
        }
        streamEnd = repaired.streamEnd;
        endstream = repaired.terminator;
      }
      const endobj = skipWhitespace(source, endstream + 9);
      if (source.slice(endobj, endobj + 6) !== "endobj") {
        fail("unsupported-document", "Stream object terminator is malformed.");
      }
      objectEnd = endobj + 6;
      encodedStreamBytes = bytes.slice(streamStart, streamEnd);
      streamBytes =
        // A chain this engine cannot decode has no decoded bytes to offer;
        // only a reader that asked for that gets here. A raw reader asked
        // for none.
        streamEncoding === "opaque" || streams === "raw"
          ? null
          : streamEncoding === "flate" ||
              streamEncoding === "flate-predictor15" ||
              streamEncoding === "flate-png-predictor"
            ? decodeFlateStream(encodedStreamBytes)
            : streamEncoding === "ascii85-flate"
              ? decodeFlateStream(decodeAscii85Stream(encodedStreamBytes))
              : streamEncoding === "ascii85-dct"
                ? decodeAscii85Stream(encodedStreamBytes)
                : streamEncoding === "ascii-hex"
                  ? decodeAsciiHexStream(encodedStreamBytes)
                  : streamEncoding === "lzw"
                    ? decodeLzwStream(encodedStreamBytes)
                    : Uint8Array.from(encodedStreamBytes);
    } else {
      const endobj = findTokenAfter(source, "endobj", afterDictionary);
      objectEnd = endobj + 6;
    }
  } else {
    const endobj = findTokenAfter(source, "endobj", bodyStart);
    objectEnd = endobj + 6;
  }

  return Object.freeze({
    objectNumber,
    generation,
    storage: "direct" as const,
    start: offset,
    end: objectEnd,
    bytes: bytes.slice(offset, objectEnd),
    dictionaryStart,
    dictionaryEnd: dictionaryStop,
    lengthNumberStart,
    lengthNumberEnd,
    ...(lengthKeyMissing ? { lengthKeyMissing: true } : {}),
    lengthReference,
    streamStart,
    streamEnd,
    streamEncoding,
    encodedStreamBytes,
    streamBytes,
  });
}

interface XrefRootFallback {
  readonly rootObjectNumber: number;
  readonly rootGeneration: number;
}

function parseClassicXref(
  bytes: Uint8Array,
  source: string,
  offset: number,
  fallbackRoot: XrefRootFallback | null = null,
): PdfXrefSection {
  exactPositiveInteger(offset, "xref offset", true);
  if (source.slice(offset, offset + 4) !== "xref") {
    fail("unsupported-document", "PDF does not use a classic xref table.");
  }
  let cursor = offset + 4;
  const entries = new Map<number, XrefEntry>();

  while (true) {
    cursor = skipWhitespace(source, cursor);
    if (source.slice(cursor, cursor + 7) === "trailer") break;
    const subsection = readLine(source, cursor);
    cursor = subsection.next;
    const header = /^(\d+)\s+(\d+)$/.exec(subsection.line);
    if (!header) {
      fail("unsupported-document", "Classic xref subsection is malformed.");
    }
    const first = Number(header[1]);
    const count = Number(header[2]);
    // Word writes an empty `0 0` subsection in front of a hybrid trailer.
    if (
      !Number.isSafeInteger(first) ||
      !Number.isSafeInteger(count) ||
      first < 0 ||
      count < 0 ||
      count > 1_000_000
    ) {
      fail("unsupported-document", "Classic xref subsection is unsupported.");
    }
    for (let index = 0; index < count; index += 1) {
      const row = readLine(source, cursor);
      cursor = row.next;
      const entry = /^(\d{10})\s(\d{5})\s([nf])(?:\s|$)/.exec(row.line);
      if (!entry) {
        fail("unsupported-document", "Classic xref entry is malformed.");
      }
      const objectNumber = first + index;
      if (entries.has(objectNumber)) {
        fail("unsupported-document", "Classic xref ownership is ambiguous.");
      }
      // Some TeX drivers write an unused slot as `0000000000 00000 n`; offset
      // 0 is the file header, never an object, so the slot is free the way a
      // reader treats it (a reference to it is simply unresolvable).
      const inUse = entry[3] === "n" && Number(entry[1]) !== 0;
      entries.set(
        objectNumber,
        Object.freeze({
          objectNumber,
          generation: Number(entry[2]),
          offset: Number(entry[1]),
          inUse,
          kind: inUse ? ("direct" as const) : ("free" as const),
          objectStreamNumber: null,
          objectStreamIndex: null,
        }),
      );
    }
  }

  cursor = skipWhitespace(source, cursor + 7);
  const trailerStart = cursor;
  const trailerEnd = dictionaryEnd(source, trailerStart);
  const trailerRaw = source.slice(trailerStart, trailerEnd);
  // A classic trailer may carry a nested dictionary (a cross-reference
  // stream dictionary copied whole, with its /DecodeParms, by Distiller);
  // only the trailer's own keys are read below, so nesting is refused only
  // when a nested dictionary could shadow one of them.
  const nestedTrailer = source.slice(trailerStart + 2, trailerEnd - 2);
  if (
    nestedTrailer.includes("<<") &&
    /<<(?:[^<>]|<<[^<>]*>>)*\/(?:Size|Prev|Root|XRefStm|Encrypt)\b(?:[^<>]|<<[^<>]*>>)*>>/.test(
      nestedTrailer.slice(nestedTrailer.indexOf("<<")),
    )
  ) {
    fail(
      "unsupported-document",
      "Nested trailer dictionaries are unsupported.",
    );
  }
  if (/\/Encrypt\b/.test(trailerRaw)) {
    fail("unsupported-document", "Encrypted PDFs are preserve-only.");
  }
  const size = oneInteger(trailerRaw, /\/Size\s+(\d+)\b/g, "trailer Size");
  const previousOffset = oneInteger(
    trailerRaw,
    /\/Prev\s+(\d+)\b/g,
    "trailer Prev",
  );
  const roots = [...trailerRaw.matchAll(/\/Root\s+(\d+)\s+(\d+)\s+R\b/g)];
  // Linearized files may write only /Size in an older trailer; the newest
  // trailer in the chain then carries the catalog authority.
  if (size === null || roots.length > 1 || (roots.length !== 1 && !fallbackRoot)) {
    fail("unsupported-document", "PDF trailer Size/Root is malformed.");
  }
  const rootObjectNumber = roots.length === 1 ? Number(roots[0][1]) : fallbackRoot!.rootObjectNumber;
  const rootGeneration = roots.length === 1 ? Number(roots[0][2]) : fallbackRoot!.rootGeneration;
  // Hybrid-reference files (PDF 1.5+, written by Word among others) keep the
  // objects living in object streams in an xref stream named by /XRefStm.
  // Its entries only fill object numbers the table itself does not use.
  const hybridOffset = oneInteger(trailerRaw, /\/XRefStm\s+(\d+)\b/g, "trailer XRefStm");
  if (hybridOffset !== null) {
    if (hybridOffset === offset) {
      fail("unsupported-document", "PDF hybrid xref stream offset is invalid.");
    }
    const hybrid = parseXrefStream(bytes, hybridOffset, { rootObjectNumber, rootGeneration }, true);
    for (const [objectNumber, entry] of hybrid.entries) {
      const existing = entries.get(objectNumber);
      if (existing === undefined || !existing.inUse) entries.set(objectNumber, entry);
    }
  }
  return Object.freeze({
    kind: "classic" as const,
    offset,
    entries,
    trailerRaw,
    size,
    rootObjectNumber,
    rootGeneration,
    previousOffset,
  });
}

function oneIntegerArray(
  dictionary: string,
  key: string,
  label: string,
): readonly number[] | null {
  const matches = [
    ...dictionary.matchAll(new RegExp(`/${key}\\s*\\[([^\\]]*)\\]`, "g")),
  ];
  if (matches.length === 0) return null;
  if (matches.length !== 1 || !/^\s*\d+(?:\s+\d+)*\s*$/.test(matches[0][1])) {
    fail("unsupported-document", `PDF ${label} is malformed or ambiguous.`);
  }
  return Object.freeze(
    matches[0][1]
      .trim()
      .split(/\s+/)
      .map((value) => Number(value)),
  );
}

function readUnsignedBigEndian(
  bytes: Uint8Array,
  start: number,
  width: number,
): number {
  let value = 0;
  for (let index = 0; index < width; index += 1) {
    value = value * 256 + bytes[start + index];
  }
  if (!Number.isSafeInteger(value)) {
    fail("unsupported-document", "PDF xref stream field exceeds safe bounds.");
  }
  return value;
}

function parseXrefStream(
  bytes: Uint8Array,
  offset: number,
  fallbackRoot: XrefRootFallback | null = null,
  hybrid = false,
): PdfXrefSection {
  const object = parseDirectObjectAtOffset(bytes, offset);
  if (
    object.dictionaryStart === null ||
    object.dictionaryEnd === null ||
    object.streamBytes === null
  ) {
    fail("unsupported-document", "PDF xref stream object is malformed.");
  }
  const source = decodedSourceLatin1(bytes);
  const dictionary = source.slice(object.dictionaryStart, object.dictionaryEnd);
  if ([...dictionary.matchAll(/\/Type\s*\/XRef\b/g)].length !== 1) {
    fail(
      "unsupported-document",
      "PDF xref stream Type is missing or ambiguous.",
    );
  }
  if (/\/Encrypt\b|\/XRefStm\b/.test(dictionary)) {
    fail(
      "unsupported-document",
      "Encrypted or hybrid xref PDFs are unsupported.",
    );
  }
  const declaredSize = oneInteger(dictionary, /\/Size\s+(\d+)\b/g, "xref Size");
  const previousOffset = oneInteger(
    dictionary,
    /\/Prev\s+(\d+)\b/g,
    "xref Prev",
  );
  const roots = [...dictionary.matchAll(/\/Root\s+(\d+)\s+(\d+)\s+R\b/g)];
  const rootObjectNumber =
    roots.length === 0 && fallbackRoot ? fallbackRoot.rootObjectNumber : Number(roots[0]?.[1]);
  const rootGeneration =
    roots.length === 0 && fallbackRoot ? fallbackRoot.rootGeneration : Number(roots[0]?.[2]);
  const widths = oneIntegerArray(dictionary, "W", "xref W");
  if (
    declaredSize === null ||
    !Number.isSafeInteger(declaredSize) ||
    declaredSize < 1 ||
    declaredSize > 9_999_999_999 ||
    roots.length > 1 ||
    (roots.length !== 1 && !fallbackRoot) ||
    !Number.isSafeInteger(rootObjectNumber) ||
    rootObjectNumber < 1 ||
    !Number.isSafeInteger(rootGeneration) ||
    rootGeneration < 0 ||
    rootGeneration > 65_535 ||
    widths === null ||
    widths.length !== 3 ||
    widths.some((width) => width < 0 || width > 6) ||
    widths[0] + widths[1] + widths[2] < 1
  ) {
    fail("unsupported-document", "PDF xref stream dictionary is unsupported.");
  }
  const indexValues = oneIntegerArray(dictionary, "Index", "xref Index") ?? [
    0,
    declaredSize,
  ];
  if (indexValues.length < 2 || indexValues.length % 2 !== 0) {
    fail("unsupported-document", "PDF xref stream Index is malformed.");
  }
  const objectNumbers: number[] = [];
  const seen = new Set<number>();
  for (let cursor = 0; cursor < indexValues.length; cursor += 2) {
    const first = indexValues[cursor];
    const count = indexValues[cursor + 1];
    if (
      !Number.isSafeInteger(first) ||
      !Number.isSafeInteger(count) ||
      first < 0 ||
      count < 1 ||
      first + count > 9_999_999_999 ||
      objectNumbers.length + count > 1_000_000
    ) {
      fail("unsupported-document", "PDF xref stream Index is unsupported.");
    }
    for (
      let objectNumber = first;
      objectNumber < first + count;
      objectNumber += 1
    ) {
      if (seen.has(objectNumber)) {
        fail("unsupported-document", "PDF xref stream ownership is ambiguous.");
      }
      seen.add(objectNumber);
      objectNumbers.push(objectNumber);
    }
  }
  // A hybrid file's /XRefStm (the IRS forms) may state a /Size below the
  // objects its /Index describes; readers take the objects as described. The
  // section is as large as the larger of the two, so a revision never numbers
  // a new object over one of them (mirror of replay).
  const size = objectNumbers.reduce(
    (largest, objectNumber) => Math.max(largest, objectNumber + 1),
    declaredSize,
  );
  const rowWidth = widths[0] + widths[1] + widths[2];
  let payload = object.streamBytes;
  if (object.streamEncoding === "flate-png-predictor") {
    const parmsEntries = topLevelDictionaryEntries(dictionary).filter(
      (entry) => entry.key === "DecodeParms",
    );
    const predictor =
      parmsEntries.length === 1
        ? xrefPredictorParameters(
            dictionary.slice(parmsEntries[0].valueStart, parmsEntries[0].valueEnd),
          )
        : null;
    if (!predictor || predictor.columns !== rowWidth) {
      fail("unsupported-document", "PDF xref stream predictor columns do not match /W.");
    }
    payload = unpredictPngRows(object.streamBytes, rowWidth);
  }
  if (payload.length !== objectNumbers.length * rowWidth) {
    fail("unsupported-document", "PDF xref stream payload length is invalid.");
  }
  const entries = new Map<number, XrefEntry>();
  let cursor = 0;
  for (const objectNumber of objectNumbers) {
    const type =
      widths[0] === 0
        ? 1
        : readUnsignedBigEndian(payload, cursor, widths[0]);
    cursor += widths[0];
    const field2 = readUnsignedBigEndian(payload, cursor, widths[1]);
    cursor += widths[1];
    const field3 = readUnsignedBigEndian(payload, cursor, widths[2]);
    cursor += widths[2];
    if (type === 0) {
      if (field3 > 65_535) {
        fail("unsupported-document", "PDF free xref generation is invalid.");
      }
      entries.set(
        objectNumber,
        Object.freeze({
          objectNumber,
          generation: field3,
          offset: field2,
          inUse: false,
          kind: "free" as const,
          objectStreamNumber: null,
          objectStreamIndex: null,
        }),
      );
    } else if (type === 1) {
      if (field3 > 65_535) {
        fail("unsupported-document", "PDF direct xref generation is invalid.");
      }
      entries.set(
        objectNumber,
        Object.freeze({
          objectNumber,
          generation: field3,
          offset: field2,
          inUse: true,
          kind: "direct" as const,
          objectStreamNumber: null,
          objectStreamIndex: null,
        }),
      );
    } else if (type === 2) {
      if (field2 < 1 || field3 < 0 || field3 >= 10_000) {
        fail(
          "unsupported-document",
          "PDF compressed xref entry is unsupported.",
        );
      }
      entries.set(
        objectNumber,
        Object.freeze({
          objectNumber,
          generation: 0,
          offset: -1,
          inUse: true,
          kind: "compressed" as const,
          objectStreamNumber: field2,
          objectStreamIndex: field3,
        }),
      );
    } else {
      fail(
        "unsupported-document",
        "PDF xref stream entry type is unsupported.",
      );
    }
  }
  // An entry that marks the stream's own number free (PDFreactor writes
  // `0 0 65535` for it) says what an absent one says: readers find the
  // stream by startxref, never through the table. It is dropped like an
  
  if (entries.get(object.objectNumber)?.kind === "free") entries.delete(object.objectNumber);
  const self = entries.get(object.objectNumber);
  // A hybrid /XRefStm stream is owned by the classic table that names it.
  // A writer may leave its cross-reference stream out of its own table (its
  // /Index ranges skip the stream's number); readers find the stream by
  // startxref and never through the table, so an absent entry says nothing.
  // One that is there has to say where the stream is -- as any entry does,
  // very nearly: a writer that miscounts its objects by a byte miscounts this
  // one too, and the repair pass then holds it to the stream's own header
  // like every other entry (repairedXrefOffsets).
  const selfMisplaced =
    self !== undefined &&
    (self.kind !== "direct" ||
      self.generation !== object.generation ||
      Math.abs(self.offset - offset) > MAX_XREF_OFFSET_DRIFT);
  if (!hybrid && selfMisplaced) {
    fail("unsupported-document", "PDF xref stream does not own itself.");
  }
  // A hybrid stream the classic table names by /XRefStm is found through
  // that name; an entry of its own that puts it elsewhere (an IRS form's says
  // offset 0) is one no reader follows, so it is not kept (mirror of replay).
  if (hybrid && selfMisplaced) entries.delete(object.objectNumber);
  return Object.freeze({
    kind: "stream" as const,
    offset,
    entries,
    trailerRaw: dictionary,
    size,
    rootObjectNumber,
    rootGeneration,
    previousOffset: hybrid ? null : previousOffset,
  });
}

/** Every cross-reference section the file holds that reads as one: each
 * `xref` keyword that opens a classic table (wherever on its line it sits),
 * and each object whose dictionary says `/Type /XRef`. */
function locatedXrefSections(
  bytes: Uint8Array,
  source: string,
  fallbackRoot: XrefRootFallback | null,
): PdfXrefSection[] {
  const located = new Map<number, PdfXrefSection>();
  for (const match of matchAllBehind(source, /xref(?=[\x00\t\n\f\r ])/g, XREF_BEHIND)) {
    try {
      located.set(match.index, parseClassicXref(bytes, source, match.index, fallbackRoot));
    } catch (error) {
      if (!(error instanceof PdfNativeTransactionError)) throw error;
    }
  }
  const headers = [...matchAllBehind(source, /\d+[\x00\t\n\f\r ]+\d+[\x00\t\n\f\r ]+obj\b/g, { char: /[0-9]/, not: true })].map(
    (match) => match.index,
  );
  let header = 0;
  for (const match of source.matchAll(/\/Type[\x00\t\n\f\r ]*\/XRef\b/g)) {
    while (header + 1 < headers.length && headers[header + 1] < match.index) header += 1;
    const offset = headers[header];
    if (offset === undefined || offset > match.index || located.has(offset)) continue;
    try {
      located.set(offset, parseXrefStream(bytes, offset, fallbackRoot));
    } catch (error) {
      if (!(error instanceof PdfNativeTransactionError)) throw error;
    }
  }
  return [...located.values()];
}

/** The section a `startxref` or `/Prev` means when the offset it records leads
 * nowhere: the newest revision the file still holds outside the chain already
 * read. That is a section no other section names, neither as its /Prev nor as
 * the /XRefStm half of a hybrid table, and of those the last one written. A
 * linearized file's first-page table names the main one, so it is the head of
 * the two however they lie in the file. Null when the file holds none. */
function recoveredXrefSection(
  bytes: Uint8Array,
  source: string,
  chainOffsets: ReadonlySet<number>,
  fallbackRoot: XrefRootFallback | null,
): PdfXrefSection | null {
  const located = locatedXrefSections(bytes, source, fallbackRoot).filter(
    (section) => !chainOffsets.has(section.offset),
  );
  const named = new Set<number>();
  for (const section of located) {
    if (section.previousOffset !== null) named.add(section.previousOffset);
    const hybrid = /\/XRefStm\s+(\d+)\b/.exec(section.trailerRaw);
    if (section.kind === "classic" && hybrid) named.add(Number(hybrid[1]));
  }
  let newest: PdfXrefSection | null = null;
  for (const section of located) {
    if (named.has(section.offset)) continue;
    if (newest === null || section.offset > newest.offset) newest = section;
  }
  return newest;
}

function parseXrefSection(
  bytes: Uint8Array,
  source: string,
  offset: number,
  fallbackRoot: XrefRootFallback | null = null,
  chainOffsets: ReadonlySet<number> = new Set(),
): PdfXrefSection {
  // An offset may land on the whitespace before its section (a linearized
  // file pads its first-page table); readers step over it to the section.
  const start = offset < source.length ? skipWhitespace(source, offset) : offset;
  if (source.slice(start, start + 4) === "xref") {
    return parseClassicXref(bytes, source, start, fallbackRoot);
  }
  try {
    return parseXrefStream(bytes, start, fallbackRoot);
  } catch (error) {
    // A writer may record a `startxref` or `/Prev` that misses its table
    // (counting from the wrong place, or from before a later edit moved it).
    // Readers find the table anyway, and so does this, only when the recorded
    // offset leads nowhere (mirror of replay).
    if (!(error instanceof PdfNativeTransactionError)) throw error;
    // The section refused may be the encrypted one, and the search passes
    // over every section it cannot read.
    refuseEncryptionNamedAnywhere(source);
    const recovered = recoveredXrefSection(bytes, source, chainOffsets, fallbackRoot);
    if (recovered === null || recovered.offset === start) throw error;
    return recovered;
  }
}

/** How far a table entry may miss its object and still be read as naming it:
 * a miscount, an end-of-line conversion, an edit the writer did not fold
 * back in. Beyond this the entry names somewhere else, not this object. */
const MAX_XREF_OFFSET_DRIFT = 64;

/** Where each object writes its own header at the start of a line, by
 * `number generation`, in file order. */
function objectHeaderOffsets(source: string): ReadonlyMap<string, readonly number[]> {
  const found = new Map<string, number[]>();
  for (const match of source.matchAll(
    /(^|[\r\n])(\d+)[\x00\t\f ]+(\d+)[\x00\t\f ]+obj\b/g,
  )) {
    const key = `${Number(match[2])} ${Number(match[3])}`;
    const offsets = found.get(key) ?? [];
    offsets.push(match.index + match[1].length);
    found.set(key, offsets);
  }
  return found;
}

/** Where an object a table entry misses really is, `"undefined"` when its
 * revision never defines it, or null when the file does not say
 * unambiguously. A table names only what its own revision had written, the
 * bytes before the `%%EOF` that closes it; a later revision's copy of the
 * object is not one it could mean. With no definition there at all the entry
 * names an object that does not exist, which a reference reads as null (PDF
 * 32000-1 7.3.10), as a classic row at offset 0 already does. Otherwise a
 * table that names very nearly the right place still names the object: the
 * one header beside the recorded offset. One that names nowhere near it still
 * names the object when that revision defines it only once, as readers take
 * it; with more than one definition and none beside the recorded offset, the
 * entry names none. */
function repairedObjectHeaderOffset(
  headers: ReadonlyMap<string, readonly number[]>,
  objectNumber: number,
  generation: number,
  recorded: number,
  revisionEnd: number,
): number | "undefined" | null {
  const all = (headers.get(`${objectNumber} ${generation}`) ?? []).filter(
    (offset) => offset < revisionEnd,
  );
  if (all.length === 0) return "undefined";
  const near = all.filter(
    (offset) => Math.abs(offset - recorded) <= MAX_XREF_OFFSET_DRIFT,
  );
  if (near.length === 1) return near[0];
  if (near.length === 0 && all.length === 1) return all[0];
  return null;
}

/** The same sections with every entry that misses its object moved onto the
 * object it names, where the file says unambiguously where that is. A table
 * whose every entry already owns its object is returned untouched. */
function repairedXrefOffsets(
  source: string,
  parsed: ParsedPdf,
): ParsedPdf {
  let repairedAny = false;
  let headers: ReadonlyMap<string, readonly number[]> | null = null;
  const relocated = new Map<number, XrefEntry>();
  const newer = new Set<number>();
  // A cross-reference stream is found by startxref and /Prev, never through a
  // table, so its own entry is read where it is but never written again.
  const streamSections = new Set(
    parsed.sections.filter((section) => section.kind === "stream").map((section) => section.offset),
  );
  const sections = parsed.sections.map((section) => {
    let sectionRepaired = false;
    const closing = source.indexOf("%%EOF", section.offset);
    const revisionEnd = closing < 0 ? source.length : closing;
    const entries = new Map<number, XrefEntry>();
    for (const [objectNumber, entry] of section.entries) {
      entries.set(objectNumber, entry);
      if (!entry.inUse || entry.kind !== "direct") continue;
      const header = new RegExp(
        `^${entry.objectNumber}\\s+${entry.generation}\\s+obj\\b`,
      );
      if (
        entry.offset >= 0 &&
        entry.offset < source.length &&
        header.test(source.slice(entry.offset))
      ) {
        continue;
      }
      headers ??= objectHeaderOffsets(source);
      const repaired = repairedObjectHeaderOffset(
        headers,
        entry.objectNumber,
        entry.generation,
        entry.offset,
        revisionEnd,
      );
      if (repaired === null) continue;
      sectionRepaired = true;
      if (repaired === "undefined") {
        entries.set(
          objectNumber,
          Object.freeze({ ...entry, offset: 0, inUse: false, kind: "free" as const }),
        );
        continue;
      }
      const moved = Object.freeze({ ...entry, offset: repaired });
      entries.set(objectNumber, moved);
      if (!newer.has(objectNumber) && !streamSections.has(repaired)) {
        relocated.set(objectNumber, moved);
      }
    }
    for (const objectNumber of section.entries.keys()) newer.add(objectNumber);
    if (!sectionRepaired) return section;
    repairedAny = true;
    return Object.freeze({ ...section, entries });
  });
  if (!repairedAny) return parsed;
  if (relocated.size > 0) {
    sections[0] = Object.freeze({ ...sections[0], relocated });
  }
  return Object.freeze({
    current: sections[0],
    previous: sections[1] ?? null,
    sections: Object.freeze(sections),
  });
}

/** A source's xref chain, validated once per buffer and reading. A save
 * parses its source to prepare the candidate and again to verify it, and the
 * planner's reads parse the same buffer before either; the chain, its
 * ownership walk and the object streams it opens are a function of the bytes
 * alone. Sources are never written in place, and each entry dies with its
 * buffer. A source that fails is parsed -- and fails -- again. */
const parsedSources = new WeakMap<Uint8Array, Fingerprinted<{ plain?: ParsedPdf; candidate?: ParsedPdf }>>();
/** The sources parsed most recently, by length and sum, with their bytes:
 * a caller that hands in a fresh view of the same bytes (the record binder
 * reads a candidate's source prefix as a new subarray on every save) finds
 * the reading it already has, once the bytes are compared whole. */
const recentParsedSources = new Map<string, { bytes: Uint8Array; readings: { plain?: ParsedPdf; candidate?: ParsedPdf } }>();
const MAX_RECENT_PARSED_SOURCES = 8;

/** Drop the parsed sources and dictionary readings kept by content (see
 * `releaseMraiNativeTextReadingCaches`): speed only, for a long-lived host. */
export function releaseMraiNativeParseCaches(): void {
  recentParsedSources.clear();
  keptDictionaryEntries.clear();
  keptNameEntries.clear();
}

function parsePdf(
  bytes: Uint8Array,
  allowCandidateRevision: boolean,
): ParsedPdf {
  if (bytes.length < FINGERPRINTED_SOURCE_MIN_BYTES) return parsePdfUncached(bytes, allowCandidateRevision);
  const fingerprint = freshSourceFingerprint(bytes);
  let readings = parsedSources.get(bytes)?.value;
  if (!readings) {
    const contentKey = `${bytes.length}:${fingerprint}`;
    const recent = recentParsedSources.get(contentKey);
    if (recent && bytesEqual(recent.bytes, bytes)) {
      readings = recent.readings;
      parsedSources.set(bytes, { fingerprint, value: readings });
    }
  }
  const known = allowCandidateRevision ? readings?.candidate : readings?.plain;
  if (known) return known;
  const parsed = parsePdfUncached(bytes, allowCandidateRevision);
  if (!readings) {
    readings = {};
    parsedSources.set(bytes, { fingerprint, value: readings });
    const contentKey = `${bytes.length}:${fingerprint}`;
    if (recentParsedSources.size >= MAX_RECENT_PARSED_SOURCES) recentParsedSources.delete(recentParsedSources.keys().next().value!);
    recentParsedSources.set(contentKey, { bytes, readings });
  }
  if (allowCandidateRevision) readings.candidate = parsed;
  else readings.plain = parsed;
  return parsed;
}

function parsePdfUncached(
  bytes: Uint8Array,
  allowCandidateRevision: boolean,
): ParsedPdf {
  const source = assertSourceEnvelope(bytes);
  let offset: number | null = finalStartXref(source);
  if (offset === null) {
    // A complete PDF followed by bytes that are not PDF is read as that
    // complete file is, the whole of it: its chain, the repairs of its
    // entries and the proof that each entry owns its object never look at
    // the bytes after it. Every offset is the same offset in `bytes`, so the
    // sections are the caller's to read `bytes` by, and a revision is
    // appended after the last byte of `bytes`, where a file has to end. Any
    // other file without an end keeps the refusal, and so does a candidate:
    // a file this engine wrote ends with its own revision (mirror of
    // replay's `_parse_pdf`).
    const complete = allowCandidateRevision ? null : completePdfLengthBeforeTrailingBytes(source);
    if (complete === null) {
      fail("unsupported-document", "PDF startxref/EOF boundary is malformed.");
    }
    refuseEncryptionNamedAnywhere(source);
    return parsePdfUncached(bytes.subarray(0, complete), false);
  }
  const sections: PdfXrefSection[] = [];
  const seenOffsets = new Set<number>();
  const chainOffsets = new Set<number>();
  // Keep cumulative native edits bounded while admitting the current source
  // produced by an earlier verified append-only transaction.  One base xref
  // section plus at most 31 incremental sections keeps ownership validation
  // finite; the next candidate is rejected before persistence once this cap is
  // reached.
  const maxSections = 32;
  while (offset !== null) {
    if (seenOffsets.has(offset) || sections.length >= maxSections) {
      fail("unsupported-document", "PDF xref revision chain is unsupported.");
    }
    seenOffsets.add(offset);
    const section = parseXrefSection(
      bytes,
      source,
      offset,
      sections.length > 0
        ? { rootObjectNumber: sections[0].rootObjectNumber, rootGeneration: sections[0].rootGeneration }
        : null,
      chainOffsets,
    );
    if (chainOffsets.has(section.offset)) {
      fail("unsupported-document", "PDF xref revision chain is unsupported.");
    }
    chainOffsets.add(section.offset);
    sections.push(section);
    offset = section.previousOffset;
  }
  const current = sections[0];
  if (allowCandidateRevision && !isAppendedRevisionShape(sections)) {
    fail(
      "unsupported-document",
      "Candidate PDF does not contain one supported incremental revision.",
    );
  }
  const parsed = repairedXrefOffsets(
    source,
    Object.freeze({
      current,
      previous: sections[1] ?? null,
      sections: Object.freeze(sections),
    }),
  );
  if (assertXrefOwnership(bytes, parsed)) {
    assertXfaRenderingAuthority(bytes, parsed);
  }
  return parsed;
}

/** Whether a candidate's newest section is the one revision this engine
 * appends (`incrementalXrefSection`): a classic table over a classic or
 * hybrid section, or -- over a source whose newest section is a
 * cross-reference stream -- one cross-reference stream linked to exactly that
 * section, listing only objects in use and itself where it stands, under the
 * number its /Size makes last and that no older revision uses, so it shadows
 * nothing. A classic table over a stream section is the shape Preview cannot
 * open at all, so it is never a candidate. */
function isAppendedRevisionShape(sections: readonly PdfXrefSection[]): boolean {
  if (sections.length < 2) return false;
  const [current, previous] = sections;
  if (current.kind === "classic") return previous.kind === "classic";
  if (previous.kind !== "stream" || current.previousOffset !== previous.offset) {
    return false;
  }
  const selfNumber = current.size - 1;
  const self = current.entries.get(selfNumber);
  return (
    self !== undefined &&
    self.kind === "direct" &&
    self.generation === 0 &&
    self.offset === current.offset &&
    [...current.entries.values()].every(
      (entry) => entry.inUse && entry.kind === "direct",
    ) &&
    sections.slice(1).every((section) => !section.entries.has(selfNumber))
  );
}

/**
 * Return the independently validated xref-chain depth for product admission.
 * A caller still needs receipt/journal authority before permitting depth > 1;
 * parsing support alone is never mutation authority.
 */
export function readMraiNativePdfXrefSectionCount(
  sourceBytes: Uint8Array,
): number {
  return parsePdf(sourceBytes, false).sections.length;
}

interface ObjectStreamMember {
  readonly objectNumber: number;
  readonly index: number;
  readonly bytes: Uint8Array;
}

interface NativeObjectReadContext {
  source: string;
  readonly members: Map<number, readonly ObjectStreamMember[]>;
}

function latestEntries(parsed: ParsedPdf): ReadonlyMap<number, XrefEntry> {
  const latest = new Map<number, XrefEntry>();
  for (const section of [...parsed.sections].reverse()) {
    for (const [objectNumber, entry] of section.entries) {
      latest.set(objectNumber, entry);
    }
  }
  return latest;
}

interface ObjectStreamContents {
  readonly members: readonly ObjectStreamMember[];
  /** Whether a member carries the interactive-form /XFA key. Producers
   * compress the catalog and its AcroForm, so this is where the key of a
   * real XFA document is found; the envelope scan never sees it. */
  readonly xfaKey: boolean;
}

function parseObjectStreamContents(
  bytes: Uint8Array,
  object: ParsedIndirectObject,
  sourceText?: string,
): ObjectStreamContents {
  if (
    object.storage !== "direct" ||
    object.dictionaryStart === null ||
    object.dictionaryEnd === null ||
    object.streamBytes === null
  ) {
    fail(
      "unsupported-document",
      "PDF object stream is not directly addressable.",
    );
  }
  const source = sourceText ?? decodedSourceLatin1(bytes);
  const dictionary = source.slice(object.dictionaryStart, object.dictionaryEnd);
  if ([...dictionary.matchAll(/\/Type\s*\/ObjStm\b/g)].length !== 1) {
    fail(
      "unsupported-document",
      "PDF object stream Type is missing or ambiguous.",
    );
  }
  const count = oneInteger(dictionary, /\/N\s+(\d+)\b/g, "object stream N");
  const first = oneInteger(
    dictionary,
    /\/First\s+(\d+)\b/g,
    "object stream First",
  );
  if (
    count === null ||
    first === null ||
    count < 1 ||
    count > 10_000 ||
    first < 1 ||
    first >= object.streamBytes.length
  ) {
    fail("unsupported-document", "PDF object stream bounds are unsupported.");
  }
  const decodedSource = asLatin1(object.streamBytes);
  const payload = decodedSource.slice(first);
  // The same rule as the envelope: encryption and a certification transform
  // are preserve-only; an approval signature or a usage-rights grant signs
  // the revision it was made over and forbids nothing.
  if (/\/Encrypt\b|\/DocMDP\b/.test(payload)) {
    fail(
      "unsupported-document",
      "Protected PDF state inside an object stream is preserve-only.",
    );
  }
  const xfaKey = /\/XFA\b/.test(payload);
  const header = decodedSource.slice(0, first);
  if (!/^\s*\d+(?:\s+\d+)+\s*$/.test(header)) {
    fail("unsupported-document", "PDF object stream header is malformed.");
  }
  const values = header.trim().split(/\s+/).map(Number);
  if (values.length !== count * 2) {
    fail("unsupported-document", "PDF object stream header count is invalid.");
  }
  const members: ObjectStreamMember[] = [];
  const seenObjects = new Set<number>();
  let previousRelativeOffset = -1;
  for (let index = 0; index < count; index += 1) {
    const objectNumber = values[index * 2];
    const relativeOffset = values[index * 2 + 1];
    const nextRelativeOffset =
      index + 1 < count
        ? values[(index + 1) * 2 + 1]
        : object.streamBytes.length - first;
    if (
      !Number.isSafeInteger(objectNumber) ||
      objectNumber < 1 ||
      seenObjects.has(objectNumber) ||
      !Number.isSafeInteger(relativeOffset) ||
      relativeOffset < 0 ||
      (index === 0 && relativeOffset !== 0) ||
      relativeOffset <= previousRelativeOffset ||
      !Number.isSafeInteger(nextRelativeOffset) ||
      nextRelativeOffset <= relativeOffset ||
      first + nextRelativeOffset > object.streamBytes.length
    ) {
      fail("unsupported-document", "PDF object stream ownership is ambiguous.");
    }
    seenObjects.add(objectNumber);
    previousRelativeOffset = relativeOffset;
    members.push(
      Object.freeze({
        objectNumber,
        index,
        bytes: object.streamBytes.slice(
          first + relativeOffset,
          first + nextRelativeOffset,
        ),
      }),
    );
  }
  return Object.freeze({ members: Object.freeze(members), xfaKey });
}

function parseObjectStreamMembers(
  bytes: Uint8Array,
  object: ParsedIndirectObject,
  sourceText?: string,
): readonly ObjectStreamMember[] {
  return parseObjectStreamContents(bytes, object, sourceText).members;
}

/** Object streams, opened once per source, reading and container. Every
 * object a save pins as untouched that sits in an object stream used to
 * inflate and decode that whole stream again (a quarterly report's save spent
 * most of its time there). The members are a function of the container's
 * bytes and of the reading that resolves its Length, so they are kept against
 * both; a member is handed out as a copy, never as the kept bytes. */
const openedObjectStreams = new WeakMap<Uint8Array, WeakMap<ParsedPdf, Map<number, ObjectStreamContents>>>();
const MAX_OPENED_OBJECT_STREAMS_PER_READING = 256;

function objectStreamContentsAt(
  bytes: Uint8Array,
  parsed: ParsedPdf,
  objectStreamNumber: number,
  containerOffset: number,
  sourceText?: string,
): ObjectStreamContents {
  let byReading = openedObjectStreams.get(bytes);
  let byOffset = byReading?.get(parsed);
  const known = byOffset?.get(containerOffset);
  if (known) return known;
  // An object stream's own Length may sit in its own object too.
  const container = parseDirectObjectAtOffset(bytes, containerOffset, {
    objectNumber: objectStreamNumber,
    generation: 0,
  }, sourceText, reference =>
    resolveStreamLengthReference(bytes, parsed, reference, sourceText),
  );
  const contents = parseObjectStreamContents(bytes, container, sourceText);
  if (!byReading) {
    byReading = new WeakMap();
    openedObjectStreams.set(bytes, byReading);
  }
  if (!byOffset) {
    byOffset = new Map();
    byReading.set(parsed, byOffset);
  }
  if (byOffset.size >= MAX_OPENED_OBJECT_STREAMS_PER_READING) byOffset.delete(byOffset.keys().next().value!);
  byOffset.set(containerOffset, contents);
  return contents;
}

/** True when the document declares an interactive-form /XFA key anywhere the
 * parser can see it: in the plain source, or inside an object stream. */
function assertXrefOwnership(bytes: Uint8Array, parsed: ParsedPdf): boolean {
  const source = decodedSourceLatin1(bytes);
  const ownedOffsets = new Map<number, string>();
  for (const section of [...parsed.sections].reverse()) {
    for (const entry of section.entries.values()) {
      if (!entry.inUse || entry.kind !== "direct") continue;
      if (entry.offset < 0 || entry.offset >= source.length) {
        fail("unsupported-document", "PDF xref object offset is out of range.");
      }
      const owner = `${entry.objectNumber}:${entry.generation}`;
      const existingOwner = ownedOffsets.get(entry.offset);
      if (existingOwner !== undefined && existingOwner !== owner) {
        fail("unsupported-document", "PDF xref object offsets are ambiguous.");
      }
      ownedOffsets.set(entry.offset, owner);
      const header = new RegExp(
        `^${entry.objectNumber}\\s+${entry.generation}\\s+obj\\b`,
      );
      if (!header.test(source.slice(entry.offset))) {
        fail(
          "unsupported-document",
          `PDF xref does not own object ${entry.objectNumber} ${entry.generation}.`,
        );
      }
    }
  }
  const latest = latestEntries(parsed);
  const objectStreamCache = new Map<number, readonly ObjectStreamMember[]>();
  const compressedOwners = new Map<string, number>();
  let xfaKey = /\/XFA\b/.test(source);
  for (const entry of latest.values()) {
    if (!entry.inUse || entry.kind !== "compressed") continue;
    const objectStreamNumber = entry.objectStreamNumber!;
    const container = latest.get(objectStreamNumber);
    if (
      container?.kind !== "direct" ||
      !container.inUse ||
      container.generation !== 0
    ) {
      fail(
        "unsupported-document",
        "PDF object stream owner does not resolve directly.",
      );
    }
    let members = objectStreamCache.get(objectStreamNumber);
    if (!members) {
      const contents = objectStreamContentsAt(bytes, parsed, objectStreamNumber, container.offset, source);
      members = contents.members;
      if (contents.xfaKey) xfaKey = true;
      objectStreamCache.set(objectStreamNumber, members);
    }
    const member = members[entry.objectStreamIndex!];
    if (!member || member.objectNumber !== entry.objectNumber) {
      fail(
        "unsupported-document",
        "PDF compressed object ownership is invalid.",
      );
    }
    const key = `${objectStreamNumber}:${entry.objectStreamIndex}`;
    const existing = compressedOwners.get(key);
    if (existing !== undefined && existing !== entry.objectNumber) {
      fail(
        "unsupported-document",
        "PDF compressed object ownership is ambiguous.",
      );
    }
    compressedOwners.set(key, entry.objectNumber);
  }
  const root = latest.get(parsed.current.rootObjectNumber);
  if (!root?.inUse || root.generation !== parsed.current.rootGeneration) {
    fail("unsupported-document", "PDF trailer Root does not resolve.");
  }
  return xfaKey;
}

/** The largest XFA packet this gate decodes to read one declaration out of it.
 * A form whose packet is larger than this is refused rather than admitted
 * without reading it. */
const MAX_XFA_INSPECTED_PACKET_BYTES = 4 * 1024 * 1024;

/** An XDP packet name as producers write it: `(config)`, `(template)`,
 * `(xdp:xdp)`, `(</xdp:xdp>)`. Printable ASCII with no escape, so the literal
 * can be compared without decoding it. */
const XFA_PACKET_NAME = /^\([\x20-\x27\x2a-\x5b\x5d-\x7e]{1,128}\)$/;

/** The document declares XFA but nothing this engine reads proves the reader
 * draws the document's own pages. Preserve-only, exactly as before. */
function refuseUnprovenXfa(): never {
  fail("unsupported-document", "XFA PDFs are preserve-only.");
}

function refuseDynamicXfa(): never {
  fail("unsupported-document", "Dynamic XFA PDFs are preserve-only.");
}

/** Read the XFA `config` packet, or null when the form ships none.
 *
 * /XFA is either one stream holding the whole XDP document, or the flattened
 * array of `(name) N G R` pairs every Designer and LiveCycle build writes. */
function readXfaConfigPacket(
  bytes: Uint8Array,
  parsed: ParsedPdf,
  acroFormSource: string,
  entry: TopLevelDictionaryEntry,
): string | null {
  const readPacket = (
    reference: Readonly<{ objectNumber: number; generation: number }>,
  ): string => {
    const packet = parseIndirectObject(
      bytes,
      parsed,
      reference.objectNumber,
      reference.generation,
      undefined,
      "opaque",
    );
    if (
      packet.streamBytes === null ||
      packet.streamBytes.length > MAX_XFA_INSPECTED_PACKET_BYTES
    ) {
      refuseUnprovenXfa();
    }
    return asLatin1(packet.streamBytes!);
  };
  if (acroFormSource[entry.valueStart] !== "[") {
    const whole = parseTopLevelReference(
      acroFormSource,
      entry.valueStart,
      entry.valueEnd,
      "PDF AcroForm XFA",
    );
    if (whole.end !== entry.valueEnd) refuseUnprovenXfa();
    return readPacket(whole.reference);
  }
  if (acroFormSource[entry.valueEnd - 1] !== "]") refuseUnprovenXfa();
  const limit = entry.valueEnd - 1;
  let cursor = entry.valueStart + 1;
  let packets = 0;
  while (true) {
    cursor = skipWhitespaceWithin(acroFormSource, cursor, limit);
    if (cursor >= limit) break;
    if (acroFormSource[cursor] !== "(") refuseUnprovenXfa();
    const nameEnd = pdfLiteralEnd(acroFormSource, cursor, limit);
    const name = acroFormSource.slice(cursor, nameEnd);
    if (!XFA_PACKET_NAME.test(name)) refuseUnprovenXfa();
    const packet = parseTopLevelReference(
      acroFormSource,
      nameEnd,
      limit,
      "PDF AcroForm XFA packet",
    );
    cursor = packet.end;
    packets += 1;
    if (packets > 64) refuseUnprovenXfa();
    if (name === "(config)") return readPacket(packet.reference);
  }
  return null;
}

/**
 * Decide whether the reader draws this document's own page content.
 *
 * An /XFA entry alone does not make a PDF preserve-only. A *static* XFA form —
 * what every Designer and LiveCycle build writes unless it is told otherwise —
 * ships real page content streams, and every reader, Acrobat included, draws
 * exactly those; the XFA template only backs the form fields. A *dynamic* form
 * is the opposite: its pages are a placeholder the reader replaces with a
 * layout it generates from the template at open time, so editing those bytes
 * edits something nobody will ever see. Two declarations say which one a file
 * is, and both are read here:
 *
 *  - the catalog's /NeedsRendering (ISO 32000-1 table 28): true means the
 *    reader regenerates the document when it opens it;
 *  - the XFA packet's own `<config> … <dynamicRender>`, whose value `required`
 *    is the form asking to be laid out at run time. Its documented default is
 *    `forbidden`, so a packet that omits it — or a form that ships no config
 *    packet at all — is static.
 *
 * Anything that cannot be read through to one of those answers is refused
 * exactly as before: the key is admitted on proof, never on missing proof.
 */
function assertXfaRenderingAuthority(
  bytes: Uint8Array,
  parsed: ParsedPdf,
): void {
  const catalog = parseIndirectObject(
    bytes,
    parsed,
    parsed.current.rootObjectNumber,
    parsed.current.rootGeneration,
  );
  const catalogSource = asLatin1(catalog.bytes);
  const needsRendering = oneTopLevelDictionaryEntry(
    catalogSource,
    "NeedsRendering",
    "PDF Catalog NeedsRendering",
    true,
  );
  if (
    needsRendering &&
    catalogSource
      .slice(needsRendering.valueStart, needsRendering.valueEnd)
      .trim() !== "false"
  ) {
    refuseDynamicXfa();
  }
  const acroFormEntry = oneTopLevelDictionaryEntry(
    catalogSource,
    "AcroForm",
    "PDF Catalog AcroForm",
    true,
  );
  // The key is only defined on the interactive form dictionary. A document
  // whose /XFA the catalog does not reach is one this engine cannot explain.
  if (!acroFormEntry) refuseUnprovenXfa();
  const acroFormSource =
    catalogSource[acroFormEntry.valueStart] === "<"
      ? catalogSource.slice(acroFormEntry.valueStart, acroFormEntry.valueEnd)
      : (() => {
          const reference = parseTopLevelReference(
            catalogSource,
            acroFormEntry.valueStart,
            acroFormEntry.valueEnd,
            "PDF Catalog AcroForm",
          );
          if (reference.end !== acroFormEntry.valueEnd) refuseUnprovenXfa();
          return asLatin1(
            parseIndirectObject(
              bytes,
              parsed,
              reference.reference.objectNumber,
              reference.reference.generation,
            ).bytes,
          );
        })();
  const xfaEntry = oneTopLevelDictionaryEntry(
    acroFormSource,
    "XFA",
    "PDF AcroForm XFA",
    true,
  );
  if (!xfaEntry) refuseUnprovenXfa();
  const config = readXfaConfigPacket(bytes, parsed, acroFormSource, xfaEntry);
  if (
    config !== null &&
    /<dynamicRender\b[^>]*>\s*required\s*<\/dynamicRender\s*>/.test(config)
  ) {
    refuseDynamicXfa();
  }
}

function resolveXrefEntry(
  parsed: ParsedPdf,
  objectNumber: number,
  generation: number,
): XrefEntry {
  let entry: XrefEntry | undefined;
  for (const section of parsed.sections) {
    entry = section.entries.get(objectNumber);
    if (entry) break;
  }
  // A reference to an object the file does not hold is refused, never read as
  // null: page numbering and resource identity past it cannot be proven. The
  // reason names which way the file is damaged (a truncated or partly zeroed
  // download typically leaves free or missing entries behind).
  if (!entry) {
    fail(
      "unsupported-document",
      `PDF object ${objectNumber} ${generation} is referenced but not in the file's cross-reference table: the file is damaged.`,
    );
  }
  if (!entry.inUse) {
    fail(
      "unsupported-document",
      `PDF object ${objectNumber} ${generation} is referenced but marked free: the file is damaged.`,
    );
  }
  if (entry.generation !== generation) {
    fail(
      "unsupported-document",
      `PDF object ${objectNumber} is referenced as generation ${generation} but the file holds generation ${entry.generation}.`,
    );
  }
  return entry;
}

/** True while a stream /Length is read out of an object stream: that object
 * stream's own Length is then read only from a plain object, so the reading
 * never recurses. */
let readingCompressedStreamLength = false;

/** Read a stream /Length a producer stored in its own object.
 *
 * The referenced object must be one plain integer, so this can never recurse
 * back into another stream. It may be a member of an object stream (as
 * Ghostscript writes every small object: each of its edited files was
 * refused as "not directly stored"), when that object stream's own Length is
 * a plain object or a number. */
function resolveStreamLengthReference(
  bytes: Uint8Array,
  parsed: ParsedPdf,
  reference: Readonly<{ objectNumber: number; generation: number }>,
  decodedSource?: string,
): number {
  const entry = resolveXrefEntry(
    parsed,
    reference.objectNumber,
    reference.generation,
  );
  let text: string;
  if (entry.kind === "compressed" && reference.generation === 0 && !readingCompressedStreamLength) {
    const containerEntry = resolveXrefEntry(parsed, entry.objectStreamNumber!, 0);
    if (containerEntry.kind !== "direct") {
      fail("unsupported-document", "Nested PDF object streams are unsupported.");
    }
    readingCompressedStreamLength = true;
    let members: ObjectStreamContents["members"];
    try {
      members = objectStreamContentsAt(
        bytes, parsed, entry.objectStreamNumber!, containerEntry.offset, decodedSource,
      ).members;
    } finally {
      readingCompressedStreamLength = false;
    }
    const member = members[entry.objectStreamIndex!];
    if (!member || member.objectNumber !== reference.objectNumber) {
      fail("unsupported-document", "PDF compressed object identity is invalid.");
    }
    text = asLatin1(member.bytes);
  } else {
    if (entry.kind !== "direct") {
      fail("unsupported-document", "PDF stream Length object is not directly stored.");
    }
    const object = parseDirectObjectAtOffset(
      bytes,
      entry.offset,
      reference,
      decodedSource,
    );
    const body = /\bobj\b([\s\S]*?)\bendobj\b/.exec(asLatin1(object.bytes));
    text = body ? body[1] : "";
  }
  const digits = /^\s*(\d{1,12})\s*$/.exec(text);
  if (!digits) {
    fail("unsupported-document", "PDF stream Length object is not one integer.");
  }
  return Number(digits[1]);
}

function parseIndirectObject(
  bytes: Uint8Array,
  parsed: ParsedPdf,
  objectNumber: number,
  generation: number,
  readContext?: NativeObjectReadContext,
  streams: "decoded" | "opaque" | "raw" = "decoded",
): ParsedIndirectObject {
  const entry = resolveXrefEntry(parsed, objectNumber, generation);
  if (entry.kind === "direct") {
    return parseDirectObjectAtOffset(bytes, entry.offset, {
      objectNumber,
      generation,
    }, readContext?.source, reference =>
      resolveStreamLengthReference(bytes, parsed, reference, readContext?.source),
      streams,
    );
  }
  if (entry.kind !== "compressed" || generation !== 0) {
    fail("unsupported-document", "PDF object ownership is unsupported.");
  }
  const containerEntry = resolveXrefEntry(parsed, entry.objectStreamNumber!, 0);
  if (containerEntry.kind !== "direct") {
    fail("unsupported-document", "Nested PDF object streams are unsupported.");
  }
  const members = objectStreamContentsAt(
    bytes, parsed, entry.objectStreamNumber!, containerEntry.offset, readContext?.source,
  ).members;
  const member = members[entry.objectStreamIndex!];
  if (!member || member.objectNumber !== objectNumber) {
    fail("unsupported-document", "PDF compressed object identity is invalid.");
  }
  return Object.freeze({
    objectNumber,
    generation: 0,
    storage: "object-stream" as const,
    start: -1,
    end: -1,
    bytes: Uint8Array.from(member.bytes),
    dictionaryStart: null,
    dictionaryEnd: null,
    lengthNumberStart: null,
    lengthNumberEnd: null,
    lengthReference: null,
    streamStart: null,
    streamEnd: null,
    streamEncoding: null,
    encodedStreamBytes: null,
    streamBytes: null,
  });
}

function assertTargetStream(
  object: ParsedIndirectObject,
): asserts object is ParsedIndirectObject & {
  lengthNumberStart: number;
  lengthNumberEnd: number;
  streamStart: number;
  streamEnd: number;
  streamEncoding: "identity" | "flate" | "ascii85-flate" | "ascii-hex" | "lzw";
  encodedStreamBytes: Uint8Array;
  streamBytes: Uint8Array;
} {
  if (
    object.storage !== "direct" ||
    object.lengthNumberStart === null ||
    object.lengthNumberEnd === null ||
    object.streamStart === null ||
    object.streamEnd === null ||
    object.streamEncoding === null ||
    object.streamEncoding === "opaque" ||
    object.streamEncoding === "dct" ||
    object.streamEncoding === "ascii85-dct" ||
    object.streamEncoding === "flate-predictor15" ||
    object.encodedStreamBytes === null ||
    object.streamBytes === null
  ) {
    fail("unsupported-document", "Target PDF object is not a content stream.");
  }
}

/** Rebuild one stream object around new encoded bytes.
 *
 * Every byte the producer wrote outside the length digits and the stream body
 * stays exactly where it was, which is what makes the revision replayable. */
type DirectlyWritableStreamObject = ParsedIndirectObject & {
  lengthNumberStart: number;
  lengthNumberEnd: number;
  streamStart: number;
  streamEnd: number;
};

/** A stream object whose /Length and stream bytes sit at known offsets of
 * the source, so its stream can be rewritten in place. */
function assertDirectlyWritableStream(
  target: ParsedIndirectObject,
): asserts target is DirectlyWritableStreamObject {
  if (
    target.lengthNumberStart === null ||
    target.lengthNumberEnd === null ||
    target.streamStart === null ||
    target.streamEnd === null
  ) {
    fail(
      "unsupported-document",
      "Native content stream is not directly writable.",
    );
  }
}

function rewriteStreamObjectBytes(
  sourceBytes: Uint8Array,
  target: DirectlyWritableStreamObject,
  encodedStream: Uint8Array,
): Uint8Array {
  return concatBytes([
    sourceBytes.slice(target.start, target.lengthNumberStart),
    // A dictionary that named no /Length gets one (mirror of replay).
    ascii(
      target.lengthKeyMissing
        ? `/Length ${encodedStream.length} `
        : String(encodedStream.length),
    ),
    sourceBytes.slice(target.lengthNumberEnd, target.streamStart),
    encodedStream,
    sourceBytes.slice(target.streamEnd, target.end),
  ]);
}

function assertObjectReference(objectNumber: number, generation: number): void {
  exactPositiveInteger(objectNumber, "PDF object number");
  exactGeneration(generation);
}

async function objectPrecondition(
  bytes: Uint8Array,
  parsed: ParsedPdf,
  objectNumber: number,
  generation: number,
  role: PdfNativeObjectPrecondition["role"],
  readContext?: NativeObjectReadContext,
): Promise<PdfNativeObjectPrecondition> {
  // Proving an object stays untouched is hashing the bytes the producer
  // wrote. A picture this engine cannot decode -- a JPX photograph, a CCITT
  // scan, an ASCIIHex hop in front of a JPEG -- is still a picture the page
  // beside it may be edited around, so the hash is taken without decoding.
  const object = parseIndirectObject(
    bytes,
    parsed,
    objectNumber,
    generation,
    readContext,
    "raw",
  );
  return Object.freeze({
    objectNumber,
    generation,
    expectedObjectSha256: await sha256PdfBytes(object.bytes),
    role,
  });
}

function pageContentReferences(
  objectSource: string,
): readonly Readonly<{ objectNumber: number; generation: number }>[] {
  // A page without /Contents is an empty page (PDF 32000-1 Table 31: "if
  // this entry is absent, the page shall be empty"). It draws nothing, and
  // only a route that needs its content refuses it (mirror of replay).
  const entry = oneTopLevelDictionaryEntry(
    objectSource,
    "Contents",
    "PDF page Contents",
    true,
  );
  if (!entry) return Object.freeze([]);
  if (objectSource[entry.valueStart] === "[") {
    return parseTopLevelReferenceArray(
      objectSource,
      entry,
      "PDF page Contents",
      1,
      32,
    );
  }
  const parsed = parseTopLevelReference(
    objectSource,
    entry.valueStart,
    entry.valueEnd,
    "PDF page Contents",
  );
  if (parsed.end !== entry.valueEnd) {
    fail("unsupported-document", "PDF page Contents are malformed.");
  }
  return Object.freeze([parsed.reference]);
}

/**
 * The streams a page draws, in order. `/Contents` names one stream, an array
 * of streams, or -- a dictionary value may be indirect (PDF 32000-1 7.3.10,
 * Table 31) -- one object holding that array, which Acrobat writes and keeps
 * in an object stream. The last is read through that object; a single
 * reference to anything but an array is left for the stream check to judge.
 * Mirrors replay's `_page_content_refs`.
 */
function resolvedPageContentReferences(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
  pageSource: string,
  readContext?: NativeObjectReadContext,
): readonly Readonly<{ objectNumber: number; generation: number }>[] {
  const references = pageContentReferences(pageSource);
  const entry = oneTopLevelDictionaryEntry(pageSource, "Contents", "PDF page Contents", true);
  if (!entry || pageSource[entry.valueStart] === "[" || references.length !== 1) {
    return references;
  }
  // Only whether the holder is a stream or an array is read: its bytes are
  // never inflated (the planner asks for a page's content once per item,
  // font and route).
  const holder = parseIndirectObject(
    sourceBytes,
    parsed,
    references[0].objectNumber,
    references[0].generation,
    readContext,
    "raw",
  );
  if (holder.streamStart !== null) return references;
  const value = asLatin1(holder.bytes)
    .replace(/^[\s\x00]*\d+\s+\d+\s+obj\b/, "")
    .replace(/\bendobj[\s\x00]*$/, "")
    .trim();
  if (!value.startsWith("[")) return references;
  return pageContentReferences(`<< /Contents ${value} >>`);
}

export function parseMraiNativePageContentReferences(
  pageObjectBytes: Uint8Array,
  sourceBytes?: Uint8Array,
): readonly Readonly<{ objectNumber: number; generation: number }>[] {
  return sourceBytes
    ? resolvedPageContentReferences(
        sourceBytes,
        // The same structure `parsePdf` would read, kept per buffer. The text
        // planner asks for this page's content once per candidate font of the
        // page and once per form it draws, and reading the xref over for each
        // of those re-decodes the whole file: on a 800 KB page drawn through
        // 29 forms with 15 font resources it was a quarter of the wait before
        // a text box opens.
        parseSourceStructure(sourceBytes),
        asLatin1(pageObjectBytes),
      )
    : pageContentReferences(asLatin1(pageObjectBytes));
}

export type PdfNativeResourceDictionaryKind =
  /** The page declares no `/Resources` at all and inherits none: it names no
   * resource, so its resource dictionary is empty. The format calls the key
   * required, and every writer of one page of lines -- pdfcpu, a plotter --
   * leaves it out. */
  | "absent-resources"
  | "nested-resources"
  | "resource-dictionary";

export interface PdfNativeReferenceDictionaryEntry {
  readonly resourceName: string;
  readonly target: Readonly<{ objectNumber: number; generation: number }>;
  /** The exact source of a resource a producer wrote as a direct dictionary
   * instead of a reference. Both spellings name the same object to a reader
   * (PDF 32000-1 7.3.8: only a stream has to be indirect), so a `/Font`
   * entry is routinely written either way. `target` then names the object
   * whose bytes this dictionary is part of, which is what pins it: any edit
   * to the dictionary changes that object's digest. */
  readonly inlineDictionary?: string;
}

export type PdfNativeExtGStateDictionaryEntry = Readonly<
  | {
      resourceName: string;
      kind: "direct";
      dictionary: string;
    }
  | {
      resourceName: string;
      kind: "indirect";
      target: Readonly<{ objectNumber: number; generation: number }>;
    }
>;

function exactDirectDictionarySource(
  source: string,
  key: string,
  label: string,
  allowMissing = false,
): string | null {
  const entry = oneTopLevelDictionaryEntry(source, key, label, allowMissing);
  if (!entry) return null;
  if (
    source.slice(entry.valueStart, entry.valueStart + 2) !== "<<" ||
    dictionaryEnd(source, entry.valueStart) !== entry.valueEnd
  ) {
    fail("unsupported-document", `${label} are not one direct dictionary.`);
  }
  return source.slice(entry.valueStart, entry.valueEnd);
}

/** An object's text as far as its top-level dictionary needs, kept per
 * buffer: a form's dictionary stands before its stream data, and a figure
 * drawn as a form carries megabytes of drawing after it. Decoding all of it
 * for every resource a plan looks up -- a chart's form read 166,657 times to
 * open one tick label -- ran for minutes (450 p10: 267s to open "10^6", 217s
 * of it in this decode). The leading bytes are read, four times as many each
 * time, until they hold the whole dictionary (an object that never closes one
 * is decoded whole, as before); what was read is kept against the buffer and
 * used again only while those same bytes sum to what they did. */
const ownerDictionarySources = new WeakMap<Uint8Array, Readonly<{ length: number; fingerprint: number; source: string }>>();
const OWNER_DICTIONARY_FIRST_READ_BYTES = 4096;

function ownerDictionarySource(ownerObjectBytes: Uint8Array): string {
  const kept = ownerDictionarySources.get(ownerObjectBytes);
  if (kept !== undefined && kept.length <= ownerObjectBytes.length &&
    bufferFingerprint(ownerObjectBytes.subarray(0, kept.length)) === kept.fingerprint) return kept.source;
  let length = Math.min(ownerObjectBytes.length, OWNER_DICTIONARY_FIRST_READ_BYTES);
  while (true) {
    const read = ownerObjectBytes.subarray(0, length);
    const source = asLatin1(read);
    let whole = length === ownerObjectBytes.length;
    if (!whole) {
      try {
        topLevelDictionaryEntries(source);
        whole = true;
      } catch (error) {
        if (!(error instanceof PdfNativeTransactionError)) throw error;
      }
    }
    if (whole) {
      ownerDictionarySources.set(ownerObjectBytes, Object.freeze({ length, fingerprint: bufferFingerprint(read), source }));
      return source;
    }
    length = Math.min(ownerObjectBytes.length, length * 4);
  }
}

function nativeResourceDictionarySource(
  ownerObjectBytes: Uint8Array,
  dictionaryKind: PdfNativeResourceDictionaryKind,
  label: string,
): string {
  if (dictionaryKind === "absent-resources") return "<<>>";
  if (dictionaryKind === "resource-dictionary") return asLatin1(ownerObjectBytes);
  const ownerSource = ownerDictionarySource(ownerObjectBytes);
  const resources = exactDirectDictionarySource(
    ownerSource,
    "Resources",
    `${label} Resources`,
  );
  if (!resources) {
    fail("unsupported-document", `${label} Resources are missing.`);
  }
  return resources;
}

export type PdfNativeResourceEntryPolicy = "strict" | "isolate-unusable";
const resourceEntryIssues = new WeakMap<object, ReadonlyMap<string, "duplicate-font-resource" | "malformed-font-resource">>();
export function readMraiNativeReferenceEntryIssue(entries: readonly PdfNativeReferenceDictionaryEntry[], name: string) {
  return resourceEntryIssues.get(entries)?.get(name) ?? null;
}

function parseReferenceDictionarySource(
  source: string,
  label: string,
  policy: PdfNativeResourceEntryPolicy = "strict",
  allowEmpty = false,
  /** The object whose bytes this dictionary is part of. Naming it admits the
   * entries a producer wrote as direct dictionaries rather than references;
   * a category that must resolve to its own object (every XObject is a
   * stream) leaves it out and keeps refusing them. */
  inlineOwner: Readonly<{ objectNumber: number; generation: number }> | null = null,
): readonly PdfNativeReferenceDictionaryEntry[] {
  // Lexical ownership must still be unambiguous. Entry-local defects do not
  // authorize choosing a duplicate or rewriting the resource dictionary.
  const entries = topLevelDictionaryEntries(source);
  const isolated = policy === "isolate-unusable";
  if (entries.length === 0 && !isolated && !allowEmpty) {
    fail("unsupported-document", `${label} is empty.`);
  }
  const counts = new Map<string, number>();
  for (const entry of entries) counts.set(entry.key, (counts.get(entry.key) ?? 0) + 1);
  const result: PdfNativeReferenceDictionaryEntry[] = [];
  const issues = new Map<string, "duplicate-font-resource" | "malformed-font-resource">();
  for (const entry of entries) {
    // A dictionary read is the page's own: its names are any regular PDF
    // names (an Adobe InDesign face /TiRoARRN~1268702012 refused every box
    
    // TAM review). The narrow form is for the names this engine makes.
    if (!PDF_EXISTING_RESOURCE_NAME.test(entry.key) || counts.get(entry.key) !== 1) {
      if (isolated) {
        issues.set(entry.key, counts.get(entry.key) !== 1 ? "duplicate-font-resource" : "malformed-font-resource");
        continue; // Exclude EVERY definition of an unusable key.
      }
      fail("unsupported-document", `${label} contains an unsupported or duplicate name /${entry.key}.`);
    }
    if (
      inlineOwner &&
      source.slice(entry.valueStart, entry.valueStart + 2) === "<<" &&
      dictionaryEnd(source, entry.valueStart) === entry.valueEnd
    ) {
      result.push(
        Object.freeze({
          resourceName: entry.key,
          target: Object.freeze({ ...inlineOwner }),
          inlineDictionary: source.slice(entry.valueStart, entry.valueEnd),
        }),
      );
      continue;
    }
    try {
      const parsed = parseTopLevelReference(source, entry.valueStart, entry.valueEnd, label);
      if (parsed.end !== entry.valueEnd) fail("unsupported-document", `${label} entry /${entry.key} is malformed.`);
      result.push(Object.freeze({ resourceName: entry.key, target: parsed.reference }));
    } catch (error) {
      if (!isolated || !(error instanceof PdfNativeTransactionError)) throw error;
      issues.set(entry.key, "malformed-font-resource");
    }
  }
  const frozen = Object.freeze(result);
  resourceEntryIssues.set(frozen, issues);
  return frozen;
}

/** The XObject entries one resource owner names, for a footprint reader:
 * an indirect /XObject dictionary is read as written through
 * `readDictionary`, since a footprint needs nothing of it that a
 * transaction would have to pin. */
export function readMraiNativeXObjectEntriesForFootprints(
  ownerObjectBytes: Uint8Array,
  dictionaryKind: PdfNativeResourceDictionaryKind,
  label: string,
  readDictionary: (ref: PdfNativeIndirectObjectRef) => string | null,
): readonly PdfNativeReferenceDictionaryEntry[] {
  const resources = nativeResourceDictionarySource(ownerObjectBytes, dictionaryKind, label);
  const entry = oneTopLevelDictionaryEntry(resources, "XObject", `${label} XObject resources`, true);
  if (!entry) return Object.freeze([]);
  if (resources.slice(entry.valueStart, entry.valueEnd).trim().startsWith("<<")) {
    return parseDirectXObjectDictionaryEntries(ownerObjectBytes, dictionaryKind, label);
  }
  const reference = oneTopLevelReference(resources, "XObject", `${label} XObject resources`);
  const dictionary = reference ? readDictionary(reference) : null;
  if (dictionary === null) {
    fail("unsupported-document", `${label} XObject resources are not readable.`);
  }
  return parseReferenceDictionarySource(dictionary, `${label} XObject resources`, "strict", true);
}

export function parseMraiNativeReferenceDictionary(
  objectBytes: Uint8Array,
  label: string,
  policy: PdfNativeResourceEntryPolicy = "strict",
  inlineOwner: Readonly<{ objectNumber: number; generation: number }> | null = null,
): readonly PdfNativeReferenceDictionaryEntry[] {
  return parseReferenceDictionarySource(
    asLatin1(objectBytes),
    label,
    policy,
    false,
    inlineOwner,
  );
}

export function parseMraiNativeIndirectFontDictionaryReference(
  ownerObjectBytes: Uint8Array,
  dictionaryKind: PdfNativeResourceDictionaryKind,
  label: string,
  allowMissing = false,
): Readonly<{ objectNumber: number; generation: number }> | null {
  const resources = nativeResourceDictionarySource(
    ownerObjectBytes,
    dictionaryKind,
    label,
  );
  const entry = oneTopLevelDictionaryEntry(
    resources,
    "Font",
    `${label} Font resources`,
    allowMissing,
  );
  if (!entry) return null;
  if (resources.slice(entry.valueStart, entry.valueStart + 2) === "<<") {
    if (dictionaryEnd(resources, entry.valueStart) !== entry.valueEnd) {
      fail("unsupported-document", `${label} Font resources are malformed.`);
    }
    return null;
  }
  const parsed = parseTopLevelReference(
    resources,
    entry.valueStart,
    entry.valueEnd,
    `${label} Font resources`,
  );
  if (parsed.end !== entry.valueEnd) {
    fail("unsupported-document", `${label} Font resources are malformed.`);
  }
  return parsed.reference;
}

export function parseMraiNativeDirectFontDictionaryEntries(
  ownerObjectBytes: Uint8Array,
  dictionaryKind: PdfNativeResourceDictionaryKind,
  label: string,
  allowMissing = false,
  policy: PdfNativeResourceEntryPolicy = "strict",
  inlineOwner: Readonly<{ objectNumber: number; generation: number }> | null = null,
): readonly PdfNativeReferenceDictionaryEntry[] {
  const resources = nativeResourceDictionarySource(
    ownerObjectBytes,
    dictionaryKind,
    label,
  );
  const fontDictionary = exactDirectDictionarySource(
    resources,
    "Font",
    `${label} Font resources`,
    allowMissing,
  );
  if (!fontDictionary) return Object.freeze([]);
  return parseReferenceDictionarySource(
    fontDictionary,
    `${label} Font resources`,
    policy,
    false,
    inlineOwner,
  );
}

/** The bytes a resource written as a direct dictionary is read through: the
 * dictionary itself, under the header of the object whose bytes hold it. Every
 * rule that reads a font reads an object, and this is the object that one is
 * part of -- so the identity and digest that pin it stay the holder's, and an
 * edit anywhere in the holder invalidates the plan. */
export function mraiNativeInlineResourceObjectBytes(
  owner: Readonly<{ objectNumber: number; generation: number }>,
  dictionary: string,
): Uint8Array {
  return latin1Bytes(
    `${owner.objectNumber} ${owner.generation} obj\n${dictionary}\nendobj`,
  );
}

/** The object a resource category is stored in when the producer wrote it as a
 * reference instead of inline.
 *
 * Which of the two spellings a producer chose carries no meaning, and the
 * planner never rewrites a resource category -- it only has to prove what the
 * names in it mean. Refusing the reference spelling refused the document. */
export function mraiNativeIndirectResourceCategory(
  ownerObjectBytes: Uint8Array,
  dictionaryKind: PdfNativeResourceDictionaryKind,
  category: string,
  label: string,
): Readonly<{ objectNumber: number; generation: number }> | null {
  const resources = nativeResourceDictionarySource(
    ownerObjectBytes,
    dictionaryKind,
    label,
  );
  const entry = oneTopLevelDictionaryEntry(
    resources,
    category,
    `${label} ${category} resources`,
    true,
  );
  if (!entry) return null;
  if (resources.slice(entry.valueStart, entry.valueStart + 2) === "<<") {
    return null;
  }
  const parsed = parseTopLevelReference(
    resources,
    entry.valueStart,
    entry.valueEnd,
    `${label} ${category} resources`,
  );
  if (parsed.end !== entry.valueEnd) {
    fail(
      "unsupported-document",
      `${label} ${category} resources are malformed.`,
    );
  }
  return parsed.reference;
}

/** Entries of one ExtGState dictionary, however the page reached it. */
export function parseMraiNativeExtGStateDictionaryObject(
  objectBytes: Uint8Array,
  label: string,
): readonly PdfNativeExtGStateDictionaryEntry[] {
  return extGStateDictionaryEntries(
    topLevelDictionarySource(asLatin1(objectBytes), `${label} ExtGState`),
    label,
  );
}

export function parseMraiNativeExtGStateDictionaryEntries(
  ownerObjectBytes: Uint8Array,
  dictionaryKind: PdfNativeResourceDictionaryKind,
  label: string,
): readonly PdfNativeExtGStateDictionaryEntry[] {
  return parseMraiNativeResourceCategoryEntries(
    ownerObjectBytes,
    dictionaryKind,
    "ExtGState",
    label,
  );
}

/**
 * The entries of one resource category, each read as the dictionary the
 * owner wrote or as the object it names. Every category a page maps by name
 * is written the same way (ISO 32000-1 7.8.3), so `/Properties` -- the
 * dictionaries a marked-content section names instead of writing inline --
 * is read exactly as `/ExtGState` is.
 */
export function parseMraiNativeResourceCategoryEntries(
  ownerObjectBytes: Uint8Array,
  dictionaryKind: PdfNativeResourceDictionaryKind,
  category: string,
  label: string,
): readonly PdfNativeExtGStateDictionaryEntry[] {
  const resources = nativeResourceDictionarySource(
    ownerObjectBytes,
    dictionaryKind,
    label,
  );
  const states = exactDirectDictionarySource(
    resources,
    category,
    `${label} ${category} resources`,
    true,
  );
  if (!states) return Object.freeze([]);
  return extGStateDictionaryEntries(states, label, category);
}

function extGStateDictionaryEntries(
  states: string,
  label: string,
  category = "ExtGState",
): readonly PdfNativeExtGStateDictionaryEntry[] {
  const names = new Set<string>();
  return Object.freeze(
    topLevelDictionaryEntries(states).map((entry) => {
      if (!PDF_NAME_TOKEN.test(entry.key)) {
        fail(
          "unsupported-document",
          `${label} ${category} resources contain an unsupported name.`,
        );
      }
      if (names.has(entry.key)) {
        fail(
          "unsupported-document",
          `${label} ${category} resources contain a duplicate name.`,
        );
      }
      names.add(entry.key);
      if (states.slice(entry.valueStart, entry.valueStart + 2) === "<<") {
        if (dictionaryEnd(states, entry.valueStart) !== entry.valueEnd) {
          fail(
            "unsupported-document",
            `${label} ${category} resource is malformed.`,
          );
        }
        return Object.freeze({
          resourceName: entry.key,
          kind: "direct" as const,
          dictionary: states.slice(entry.valueStart, entry.valueEnd),
        });
      }
      const parsed = parseTopLevelReference(
        states,
        entry.valueStart,
        entry.valueEnd,
        `${label} ${category} resource`,
      );
      if (parsed.end !== entry.valueEnd) {
        fail(
          "unsupported-document",
          `${label} ${category} resource is malformed.`,
        );
      }
      return Object.freeze({
        resourceName: entry.key,
        kind: "indirect" as const,
        target: parsed.reference,
      });
    }),
  );
}

function topLevelDictionarySource(objectSource: string, label: string): string {
  const header =
    /^\d+[\x00\x09\x0a\x0c\x0d\x20]+\d+[\x00\x09\x0a\x0c\x0d\x20]+obj\b/.exec(
      objectSource,
    );
  const start = objectSource.startsWith("<<")
    ? 0
    : skipWhitespace(objectSource, header?.[0].length ?? 0);
  if (objectSource.slice(start, start + 2) !== "<<") {
    fail("unsupported-document", `${label} dictionary is missing.`);
  }
  return objectSource.slice(start, dictionaryEnd(objectSource, start));
}

function latin1Bytes(value: string): Uint8Array {
  const result = new Uint8Array(value.length);
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    if (code > 0xff) {
      fail("unsupported-document", "PDF source contains a non-Latin-1 value.");
    }
    result[index] = code;
  }
  return result;
}

function exactDirectPdfName(
  source: string,
  start: number,
  end: number,
): string | null {
  if (source[start] !== "/") return null;
  try {
    const nameEnd = pdfNameEnd(source, start, end);
    return nameEnd === end ? decodePdfName(source, start, nameEnd) : null;
  } catch {
    return null;
  }
}

function exactBlendMode(
  dictionary: string,
  entry: TopLevelDictionaryEntry,
): PdfStandardBlendMode | null {
  const direct = exactDirectPdfName(
    dictionary,
    entry.valueStart,
    entry.valueEnd,
  );
  if (direct) return PDF_BLEND_MODE_BY_PDF_NAME[direct] ?? null;
  if (
    dictionary[entry.valueStart] !== "[" ||
    dictionary[entry.valueEnd - 1] !== "]"
  ) {
    return null;
  }
  const modes: PdfStandardBlendMode[] = [];
  let nameCount = 0;
  let cursor = entry.valueStart + 1;
  const limit = entry.valueEnd - 1;
  while (true) {
    cursor = skipWhitespace(dictionary, cursor);
    if (cursor >= limit) break;
    if (dictionary[cursor] !== "/" || nameCount >= 16) return null;
    let end: number;
    try {
      end = pdfNameEnd(dictionary, cursor, limit);
    } catch {
      return null;
    }
    const mode =
      PDF_BLEND_MODE_BY_PDF_NAME[decodePdfName(dictionary, cursor, end)];
    if (mode) modes.push(mode);
    nameCount += 1;
    cursor = end;
  }
  return modes[0] ?? null;
}

function exactExtGStateAppearance(dictionary: string): Readonly<{
  fillOpacity: number | null;
  strokeOpacity: number | null;
  blendMode: PdfStandardBlendMode | null;
}> | null {
  let entries: readonly TopLevelDictionaryEntry[];
  try {
    entries = topLevelDictionaryEntries(dictionary);
  } catch {
    return null;
  }
  const allowed = new Set(["Type", "ca", "CA", "BM", "SMask"]);
  const keys = new Set<string>();
  if (
    entries.some((entry) => {
      if (!allowed.has(entry.key) || keys.has(entry.key)) return true;
      keys.add(entry.key);
      return false;
    })
  ) {
    return null;
  }
  const type = entries.find((entry) => entry.key === "Type");
  if (type) {
    if (
      exactDirectPdfName(dictionary, type.valueStart, type.valueEnd) !==
      "ExtGState"
    ) {
      return null;
    }
  }
  const softMask = entries.find((entry) => entry.key === "SMask");
  if (
    softMask &&
    exactDirectPdfName(dictionary, softMask.valueStart, softMask.valueEnd) !==
      "None"
  ) {
    return null;
  }
  const opacity = (key: "ca" | "CA"): number | null => {
    const entry = entries.find((candidate) => candidate.key === key);
    if (!entry) return null;
    const raw = dictionary.slice(entry.valueStart, entry.valueEnd);
    if (!/^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(raw)) return null;
    const value = Number(raw);
    return Number.isFinite(value) && value >= 0 && value <= 1 ? value : null;
  };
  const fillOpacity = opacity("ca");
  const strokeOpacity = opacity("CA");
  const blendEntry = entries.find((entry) => entry.key === "BM");
  const blendMode = blendEntry
    ? exactBlendMode(dictionary, blendEntry)
    : null;
  const invalidOpacity = (key: "ca" | "CA") => {
    const entry = entries.find((candidate) => candidate.key === key);
    return Boolean(entry && opacity(key) === null);
  };
  return invalidOpacity("ca") ||
    invalidOpacity("CA") ||
    (blendEntry !== undefined && blendMode === null)
    ? null
    : Object.freeze({ fillOpacity, strokeOpacity, blendMode });
}

interface ResolvedPageExtGStateGraph {
  readonly inspection: PdfNativePageExtGStateInspection;
  readonly pageObject: ParsedIndirectObject;
  readonly resourceOwnerObject: ParsedIndirectObject;
  readonly resourcesDictionary: string;
  readonly categoryDictionary: string;
}

/** The exact slice one top-level dictionary entry occupies, with its lead. */
function dictionaryEntrySlice(
  source: string,
  key: string,
  label: string,
): Readonly<{ start: number; end: number }> {
  const matches = topLevelDictionaryEntries(source).filter(
    (entry) => entry.key === key,
  );
  if (matches.length !== 1) {
    fail("unsupported-document", `${label} is ambiguous.`);
  }
  let start = matches[0].keyStart;
  while (start > 0 && isPdfWhitespace(source[start - 1])) start -= 1;
  return Object.freeze({ start, end: matches[0].valueEnd });
}

interface ResolvedPageXObjectGraph {
  readonly pageObject: ParsedIndirectObject;
  readonly resourceOwnerObject: ParsedIndirectObject;
  /** The `/XObject` dictionary as it appears inside the owner object. */
  readonly categoryStart: number;
  readonly categoryEnd: number;
  readonly ownerSource: string;
}

/**
 * Locate the page's `/XObject` dictionary inside the object that owns it, so
 * one resource name can be dropped without touching any other byte.
 */
async function resolvePageXObjectGraph(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
  page: PdfNativePageContentGraphInspection,
): Promise<ResolvedPageXObjectGraph> {
  const pageObject = parseIndirectObject(
    sourceBytes,
    parsed,
    page.page.objectNumber,
    page.page.generation,
  );
  const resourceOwnerObject = parseIndirectObject(
    sourceBytes,
    parsed,
    page.resourceAuthority.owner.objectNumber,
    page.resourceAuthority.owner.generation,
  );
  if (
    pageObject.streamBytes !== null ||
    resourceOwnerObject.streamBytes !== null
  ) {
    fail(
      "unsupported-document",
      "PDF page resource ownership is not directly addressable.",
    );
  }
  const ownerSource = asLatin1(resourceOwnerObject.bytes);
  const resourcesDictionary =
    page.resourceAuthority.dictionaryKind === "resource-dictionary"
      ? topLevelDictionarySource(ownerSource, "PDF page Resources")
      : (exactDirectDictionarySource(
          ownerSource,
          "Resources",
          "PDF page Resources",
        ) ?? fail("unsupported-document", "PDF page Resources are missing."));
  const offset = ownerSource.indexOf(resourcesDictionary);
  if (offset < 0 || ownerSource.indexOf(resourcesDictionary, offset + 1) >= 0) {
    fail("unsupported-document", "PDF page Resources are not addressable.");
  }
  const categoryEntry = oneTopLevelDictionaryEntry(
    resourcesDictionary,
    "XObject",
    "PDF page XObject resources",
    true,
  );
  if (
    !categoryEntry ||
    resourcesDictionary.slice(
      categoryEntry.valueStart,
      categoryEntry.valueStart + 2,
    ) !== "<<"
  ) {
    // An indirect `/XObject` object may be shared with other pages, so this
    // writer leaves it alone rather than rewriting what it does not own.
    fail(
      "unsupported-document",
      "PDF page XObject resources are shared or indirect.",
    );
  }
  if (
    dictionaryEnd(resourcesDictionary, categoryEntry.valueStart) !==
    categoryEntry.valueEnd
  ) {
    fail("unsupported-document", "PDF page XObject resources are malformed.");
  }
  return Object.freeze({
    pageObject,
    resourceOwnerObject,
    categoryStart: offset + categoryEntry.valueStart,
    categoryEnd: offset + categoryEntry.valueEnd,
    ownerSource,
  });
}

export interface PdfNativePageXObjectUnmapPlan {
  readonly pageNumber: number;
  readonly page: Readonly<{ objectNumber: number; generation: number }>;
  readonly resourceOwner: Readonly<{ objectNumber: number; generation: number }>;
  readonly resourceName: string;
  readonly direction: "unmap" | "remap";
  readonly expectedOwnerSha256: string;
  /** The owner object exactly as this revision rewrites it. */
  readonly ownerBytes: Uint8Array;
  readonly revisedOwnerSha256: string;
  /** Byte offset of the entry inside the owner object, and its exact text. */
  readonly entryOffset: number;
  readonly entryBytes: Uint8Array;
}

/**
 * Drop (or put back) one `/XObject` name in the page's own resource
 * dictionary. Only a dictionary this page alone owns is rewritten, so no
 * other page can lose a resource it still paints.
 */
export async function planMraiNativePageXObjectUnmap(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly resourceName: string;
  readonly direction: "unmap" | "remap";
  /** Remap only: the exact entry slice the forward unmap removed. */
  readonly entryOffset?: number;
  readonly entryBytes?: Uint8Array;
}): Promise<PdfNativePageXObjectUnmapPlan> {
  if (!PDF_RESOURCE_NAME.test(input.resourceName)) {
    fail("invalid-request", "PDF XObject resource name is invalid.");
  }
  const parsed = parsePdf(input.sourceBytes, false);
  const pages = await inspectMraiNativeDocumentPageContentGraphs(
    input.sourceBytes,
  );
  const page = pages[input.pageNumber - 1];
  if (!page || page.pageNumber !== input.pageNumber) {
    fail("invalid-request", "Selected PDF page is out of range.");
  }
  const owner = page.resourceAuthority.owner;
  if (
    pages.some(
      (other) =>
        other.pageNumber !== page.pageNumber &&
        sameNativeObjectRef(other.resourceAuthority.owner, owner),
    )
  ) {
    fail(
      "unsupported-document",
      "PDF page resources are shared with another page.",
    );
  }
  const graph = await resolvePageXObjectGraph(input.sourceBytes, parsed, page);
  const categoryDictionary = graph.ownerSource.slice(
    graph.categoryStart,
    graph.categoryEnd,
  );
  const present = topLevelDictionaryEntries(categoryDictionary).filter(
    (entry) => entry.key === input.resourceName,
  );
  if (input.direction === "unmap") {
    if (present.length !== 1) {
      fail("precondition-mismatch", "PDF XObject resource name is not mapped.");
    }
  } else if (present.length !== 0) {
    fail("precondition-mismatch", "PDF XObject resource name is still mapped.");
  }
  // The unmap records the exact slice it cuts, so putting the name back is a
  // byte-for-byte reinsertion rather than a freshly formatted entry.
  const slice =
    input.direction === "unmap"
      ? dictionaryEntrySlice(
          categoryDictionary,
          input.resourceName,
          "PDF XObject resource name",
        )
      : null;
  const entryOffset = slice ? slice.start : (input.entryOffset ?? -1);
  const entryText = slice
    ? categoryDictionary.slice(slice.start, slice.end)
    : asLatin1(input.entryBytes ?? new Uint8Array());
  if (
    entryOffset < 0 ||
    entryOffset > categoryDictionary.length ||
    entryText.length < 2
  ) {
    fail("invalid-request", "PDF XObject resource entry authority is invalid.");
  }
  const rewrittenCategory = slice
    ? `${categoryDictionary.slice(0, slice.start)}${categoryDictionary.slice(slice.end)}`
    : `${categoryDictionary.slice(0, entryOffset)}${entryText}${categoryDictionary.slice(entryOffset)}`;
  if (
    topLevelDictionaryEntries(rewrittenCategory).filter(
      (entry) => entry.key === input.resourceName,
    ).length !== (slice ? 0 : 1)
  ) {
    fail("invalid-request", "PDF XObject resource entry rewrite is invalid.");
  }
  const ownerValue = `${graph.ownerSource.slice(0, graph.categoryStart)}${rewrittenCategory}${graph.ownerSource.slice(graph.categoryEnd)}`;
  // The bytes of an owner that lives inside an object stream are the member
  // alone. The revision writes it out as a direct object, so it carries the
  // `N G obj ... endobj` envelope back -- exactly as the page ExtGState and
  // content consolidation revisions materialize a compressed page -- and the
  // recorded hash is over the object a reader will read back.
  const ownerBytes = latin1Bytes(
    graph.resourceOwnerObject.storage === "direct"
      ? ownerValue
      : `${owner.objectNumber} ${owner.generation} obj\n${ownerValue}\nendobj`,
  );
  return Object.freeze({
    pageNumber: page.pageNumber,
    page: Object.freeze({ ...page.page }),
    resourceOwner: Object.freeze({ ...owner }),
    resourceName: input.resourceName,
    direction: input.direction,
    expectedOwnerSha256: await sha256PdfBytes(graph.resourceOwnerObject.bytes),
    ownerBytes,
    revisedOwnerSha256: await sha256PdfBytes(ownerBytes),
    entryOffset,
    entryBytes: latin1Bytes(entryText),
  });
}

async function resolvePageExtGStateGraph(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
  page: PdfNativePageContentGraphInspection,
): Promise<ResolvedPageExtGStateGraph> {
  const pageObject = parseIndirectObject(
    sourceBytes,
    parsed,
    page.page.objectNumber,
    page.page.generation,
  );
  const resourceOwnerObject = parseIndirectObject(
    sourceBytes,
    parsed,
    page.resourceAuthority.owner.objectNumber,
    page.resourceAuthority.owner.generation,
  );
  if (
    pageObject.streamBytes !== null ||
    resourceOwnerObject.streamBytes !== null
  ) {
    fail(
      "unsupported-document",
      "PDF page resource ownership is not directly addressable.",
    );
  }
  const resourcesDictionary =
    page.resourceAuthority.dictionaryKind === "resource-dictionary"
      ? topLevelDictionarySource(
          asLatin1(resourceOwnerObject.bytes),
          "PDF page Resources",
        )
      : (exactDirectDictionarySource(
          asLatin1(resourceOwnerObject.bytes),
          "Resources",
          "PDF page Resources",
        ) ?? fail("unsupported-document", "PDF page Resources are missing."));
  const categoryEntry = oneTopLevelDictionaryEntry(
    resourcesDictionary,
    "ExtGState",
    "PDF page ExtGState resources",
    true,
  );
  let categoryDictionary = "<<>>";
  let category: PdfNativePageExtGStateInspection["category"];
  if (!categoryEntry) {
    category = Object.freeze({
      kind: "missing",
      expectedDictionarySha256: await sha256PdfBytes(
        latin1Bytes(categoryDictionary),
      ),
    });
  } else if (
    resourcesDictionary.slice(
      categoryEntry.valueStart,
      categoryEntry.valueStart + 2,
    ) === "<<"
  ) {
    if (
      dictionaryEnd(resourcesDictionary, categoryEntry.valueStart) !==
      categoryEntry.valueEnd
    ) {
      fail(
        "unsupported-document",
        "PDF page ExtGState resources are malformed.",
      );
    }
    categoryDictionary = resourcesDictionary.slice(
      categoryEntry.valueStart,
      categoryEntry.valueEnd,
    );
    category = Object.freeze({
      kind: "direct",
      expectedDictionarySha256: await sha256PdfBytes(
        latin1Bytes(categoryDictionary),
      ),
    });
  } else {
    const reference = parseTopLevelReference(
      resourcesDictionary,
      categoryEntry.valueStart,
      categoryEntry.valueEnd,
      "PDF page ExtGState resources",
    );
    if (reference.end !== categoryEntry.valueEnd) {
      fail(
        "unsupported-document",
        "PDF page ExtGState resources are malformed.",
      );
    }
    const object = parseIndirectObject(
      sourceBytes,
      parsed,
      reference.reference.objectNumber,
      reference.reference.generation,
    );
    if (object.streamBytes !== null) {
      fail(
        "unsupported-document",
        "Stream-backed page ExtGState resources are preserve-only.",
      );
    }
    categoryDictionary = topLevelDictionarySource(
      asLatin1(object.bytes),
      "PDF page ExtGState resources",
    );
    category = Object.freeze({
      kind: "indirect",
      expectedDictionarySha256: await sha256PdfBytes(
        latin1Bytes(categoryDictionary),
      ),
      target: Object.freeze({
        ...reference.reference,
        expectedObjectSha256: await sha256PdfBytes(object.bytes),
      }),
    });
  }
  const names = new Set<string>();
  const entries = await Promise.all(
    topLevelDictionaryEntries(categoryDictionary).map(async (entry) => {
      if (!PDF_EXISTING_RESOURCE_NAME.test(entry.key) || names.has(entry.key)) {
        fail(
          "unsupported-document",
          "PDF page ExtGState resource identity is invalid.",
        );
      }
      names.add(entry.key);
      let dictionary: string;
      let streamBacked = false;
      if (
        categoryDictionary.slice(entry.valueStart, entry.valueStart + 2) ===
        "<<"
      ) {
        if (
          dictionaryEnd(categoryDictionary, entry.valueStart) !== entry.valueEnd
        ) {
          fail("unsupported-document", "PDF ExtGState resource is malformed.");
        }
        dictionary = categoryDictionary.slice(entry.valueStart, entry.valueEnd);
      } else {
        const reference = parseTopLevelReference(
          categoryDictionary,
          entry.valueStart,
          entry.valueEnd,
          "PDF ExtGState resource",
        );
        if (reference.end !== entry.valueEnd) {
          fail("unsupported-document", "PDF ExtGState resource is malformed.");
        }
        const object = parseIndirectObject(
          sourceBytes,
          parsed,
          reference.reference.objectNumber,
          reference.reference.generation,
        );
        streamBacked = object.streamBytes !== null;
        dictionary = topLevelDictionarySource(
          asLatin1(object.bytes),
          "PDF ExtGState resource",
        );
      }
      const appearance = streamBacked
        ? null
        : exactExtGStateAppearance(dictionary);
      return Object.freeze({
        resourceName: entry.key,
        fillOpacity: appearance?.fillOpacity ?? null,
        strokeOpacity: appearance?.strokeOpacity ?? null,
        blendMode: appearance?.blendMode ?? null,
        supportedAppearanceOnly: appearance !== null,
        dictionarySha256: await sha256PdfBytes(latin1Bytes(dictionary)),
      });
    }),
  );
  return Object.freeze({
    inspection: Object.freeze({
      sourceSha256: page.sourceSha256,
      pageNumber: page.pageNumber,
      page: Object.freeze({
        ...page.page,
        expectedObjectSha256: await sha256PdfBytes(pageObject.bytes),
      }),
      resourceOwner: Object.freeze({
        ...page.resourceAuthority.owner,
        expectedObjectSha256: await sha256PdfBytes(resourceOwnerObject.bytes),
        dictionaryKind: page.resourceAuthority.dictionaryKind,
      }),
      category,
      entries: Object.freeze(entries),
    }),
    pageObject,
    resourceOwnerObject,
    resourcesDictionary,
    categoryDictionary,
  });
}

interface ResolvedPageFontGraph {
  readonly pageObject: ParsedIndirectObject;
  readonly resourceOwnerObject: ParsedIndirectObject;
  readonly resourcesDictionary: string;
  readonly categoryDictionary: string;
  readonly category: PdfNativePageExtGStateInspection["category"];
  readonly entries: readonly PdfNativeReferenceDictionaryEntry[];
  /** Every name the category dictionary writes, including ones a writer
   * repeated or spelled badly: a new name must not be any of them. */
  readonly names: ReadonlySet<string>;
}

async function resolvePageFontGraph(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
  page: PdfNativePageContentGraphInspection,
): Promise<ResolvedPageFontGraph> {
  return resolvePageResourceCategoryGraph(sourceBytes, parsed, page, "Font");
}

async function resolvePageResourceCategoryGraph(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
  page: PdfNativePageContentGraphInspection,
  categoryName: PdfNativeContentAppendCategory,
): Promise<ResolvedPageFontGraph> {
  const pageObject = parseIndirectObject(
    sourceBytes,
    parsed,
    page.page.objectNumber,
    page.page.generation,
  );
  const resourceOwnerObject = parseIndirectObject(
    sourceBytes,
    parsed,
    page.resourceAuthority.owner.objectNumber,
    page.resourceAuthority.owner.generation,
  );
  if (
    pageObject.streamBytes !== null ||
    resourceOwnerObject.streamBytes !== null
  ) {
    fail(
      "unsupported-document",
      "PDF page font resource ownership is not directly addressable.",
    );
  }
  const resourcesDictionary =
    page.resourceAuthority.dictionaryKind === "resource-dictionary"
      ? topLevelDictionarySource(
          asLatin1(resourceOwnerObject.bytes),
          "PDF page Resources",
        )
      : (exactDirectDictionarySource(
          asLatin1(resourceOwnerObject.bytes),
          "Resources",
          "PDF page Resources",
        ) ?? fail("unsupported-document", "PDF page Resources are missing."));
  return resolveCategoryGraphFromResources(
    sourceBytes,
    parsed,
    categoryName,
    pageObject,
    resourceOwnerObject,
    resourcesDictionary,
  );
}

/** The category graph of one resources dictionary, whoever owns it: a
 * page (through its resource authority) or a Form XObject. */
async function resolveCategoryGraphFromResources(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
  categoryName: PdfNativeContentAppendCategory,
  pageObject: ParsedIndirectObject,
  resourceOwnerObject: ParsedIndirectObject,
  resourcesDictionary: string,
): Promise<ResolvedPageFontGraph> {
  const categoryEntry = oneTopLevelDictionaryEntry(
    resourcesDictionary,
    categoryName,
    `PDF page ${categoryName} resources`,
    true,
  );
  let categoryDictionary = "<<>>";
  let category: PdfNativePageExtGStateInspection["category"];
  if (!categoryEntry) {
    category = Object.freeze({
      kind: "missing" as const,
      expectedDictionarySha256: await sha256PdfBytes(
        latin1Bytes(categoryDictionary),
      ),
    });
  } else if (
    resourcesDictionary.slice(
      categoryEntry.valueStart,
      categoryEntry.valueStart + 2,
    ) === "<<"
  ) {
    if (
      dictionaryEnd(resourcesDictionary, categoryEntry.valueStart) !==
      categoryEntry.valueEnd
    ) {
      fail("unsupported-document", `PDF page ${categoryName} resources are malformed.`);
    }
    categoryDictionary = resourcesDictionary.slice(
      categoryEntry.valueStart,
      categoryEntry.valueEnd,
    );
    category = Object.freeze({
      kind: "direct" as const,
      expectedDictionarySha256: await sha256PdfBytes(
        latin1Bytes(categoryDictionary),
      ),
    });
  } else {
    const reference = parseTopLevelReference(
      resourcesDictionary,
      categoryEntry.valueStart,
      categoryEntry.valueEnd,
      `PDF page ${categoryName} resources`,
    );
    if (reference.end !== categoryEntry.valueEnd) {
      fail("unsupported-document", `PDF page ${categoryName} resources are malformed.`);
    }
    const object = parseIndirectObject(
      sourceBytes,
      parsed,
      reference.reference.objectNumber,
      reference.reference.generation,
    );
    if (object.streamBytes !== null) {
      fail(
        "unsupported-document",
        `Stream-backed page ${categoryName} resources are preserve-only.`,
      );
    }
    categoryDictionary = topLevelDictionarySource(
      asLatin1(object.bytes),
      `PDF page ${categoryName} resources`,
    );
    category = Object.freeze({
      kind: "indirect" as const,
      expectedDictionarySha256: await sha256PdfBytes(
        latin1Bytes(categoryDictionary),
      ),
      target: Object.freeze({
        ...reference.reference,
        expectedObjectSha256: await sha256PdfBytes(object.bytes),
      }),
    });
  }
  // A writer that repeats a name (FPDF writes /F5 twice) leaves that name
  // unusable, not the page: the usable entries are read around it, and a
  // new name is checked against every name written.
  const entries =
    categoryDictionary === "<<>>"
      ? Object.freeze([])
      : parseReferenceDictionarySource(
          categoryDictionary,
          `PDF page ${categoryName} resources`,
          "isolate-unusable",
        );
  const names = new Set(
    categoryDictionary === "<<>>"
      ? []
      : topLevelDictionaryEntries(categoryDictionary).map((entry) => entry.key),
  );
  return Object.freeze({
    pageObject,
    resourceOwnerObject,
    resourcesDictionary,
    categoryDictionary,
    category,
    entries,
    names,
  });
}

/** The dictionary of a stream object: the text between its header and the
 * `stream` keyword. */
function streamObjectDictionarySpan(
  source: string,
  label: string,
): Readonly<{ start: number; end: number }> {
  const header = /^[\s\x00]*\d+\s+\d+\s+obj\b/.exec(source);
  if (!header) fail("unsupported-document", `${label} header is malformed.`);
  const start = source.indexOf("<<", header[0].length);
  const streamAt = source.search(/\bstream(?:\r\n|\r|\n)/);
  if (start < 0 || streamAt < 0 || start > streamAt) {
    fail("unsupported-document", `${label} dictionary is malformed.`);
  }
  const end = dictionaryEnd(source, start);
  if (end > streamAt) fail("unsupported-document", `${label} dictionary is malformed.`);
  return Object.freeze({ start, end });
}

/** Font resources of one Form XObject, read the way a page's are: its own
 * /Resources written inline or as one indirect dictionary. */
async function resolveFormFontGraph(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
  form: PdfNativeIndirectObjectRef,
): Promise<ResolvedPageFontGraph & { readonly resourcesIndirect: ParsedIndirectObject | null }> {
  const formObject = parseIndirectObject(sourceBytes, parsed, form.objectNumber, form.generation);
  if (formObject.streamBytes === null) {
    fail("unsupported-document", "Form font resource owner is not a stream object.");
  }
  const source = asLatin1(formObject.bytes);
  const span = streamObjectDictionarySpan(source, "Form XObject");
  const dictionary = source.slice(span.start, span.end);
  if (
    [...dictionary.matchAll(/\/Type\s*\/XObject\b/g)].length !== 1 ||
    [...dictionary.matchAll(/\/Subtype\s*\/Form\b/g)].length !== 1
  ) {
    fail("unsupported-document", "Form font resource owner is not a Form XObject.");
  }
  const entry = oneTopLevelDictionaryEntry(dictionary, "Resources", "Form Resources", true);
  if (!entry) fail("unsupported-document", "Form Resources are missing.");
  if (dictionary.slice(entry.valueStart, entry.valueStart + 2) === "<<") {
    if (dictionaryEnd(dictionary, entry.valueStart) !== entry.valueEnd) {
      fail("unsupported-document", "Form Resources are malformed.");
    }
    const graph = await resolveCategoryGraphFromResources(
      sourceBytes, parsed, "Font", formObject, formObject,
      dictionary.slice(entry.valueStart, entry.valueEnd),
    );
    return Object.freeze({ ...graph, resourcesIndirect: null });
  }
  const reference = parseTopLevelReference(
    dictionary, entry.valueStart, entry.valueEnd, "Form Resources",
  );
  if (reference.end !== entry.valueEnd) {
    fail("unsupported-document", "Form Resources are malformed.");
  }
  const resourcesObject = parseIndirectObject(
    sourceBytes, parsed, reference.reference.objectNumber, reference.reference.generation,
  );
  if (resourcesObject.streamBytes !== null) {
    fail("unsupported-document", "Stream-backed Form Resources are preserve-only.");
  }
  const graph = await resolveCategoryGraphFromResources(
    sourceBytes, parsed, "Font", formObject, resourcesObject,
    topLevelDictionarySource(asLatin1(resourcesObject.bytes), "Form Resources"),
  );
  return Object.freeze({ ...graph, resourcesIndirect: resourcesObject });
}

/** Rewrite the /Resources entry of one stream object's dictionary; the
 * stream and its /Length stay exactly as written. */
function rewriteStreamObjectResources(objectBytes: Uint8Array, value: string): Uint8Array {
  const source = asLatin1(objectBytes);
  const span = streamObjectDictionarySpan(source, "Form XObject");
  const dictionary = rewriteDictionaryEntry(
    source.slice(span.start, span.end), "Resources", value, "Form Resources",
  );
  return latin1Bytes(source.slice(0, span.start) + dictionary + source.slice(span.end));
}

/** The XObject names an owner maps, its `/XObject` dictionary read through
 * the object it lives in when the owner names one (Ghostscript writes every
 * resource category that way). Reading it is all a paint needs: an edit that
 * would rewrite the dictionary refuses a shared or indirect one on its own.
 * Mirrors replay's `_xobject_entries_read_through`. */
function xObjectDictionaryEntries(
  sourceBytes: Uint8Array,
  ownerObjectBytes: Uint8Array,
  dictionaryKind: PdfNativeResourceDictionaryKind,
  label: string,
  allowMissing = false,
): readonly PdfNativeReferenceDictionaryEntry[] {
  const resources = nativeResourceDictionarySource(ownerObjectBytes, dictionaryKind, label);
  const entry = oneTopLevelDictionaryEntry(resources, "XObject", `${label} XObject resources`, true);
  if (!entry || resources.slice(entry.valueStart, entry.valueEnd).trim().startsWith("<<")) {
    return parseDirectXObjectDictionaryEntries(ownerObjectBytes, dictionaryKind, label, allowMissing);
  }
  const reference = oneTopLevelReference(resources, "XObject", `${label} XObject resources`);
  const dictionary = reference ? readMraiNativeObjectDictionarySource(sourceBytes, reference) : null;
  if (dictionary === null) {
    fail("unsupported-document", `${label} XObject resources are not readable.`);
  }
  return parseReferenceDictionarySource(dictionary, `${label} XObject resources`, "strict", true);
}

function parseDirectXObjectDictionaryEntries(
  ownerObjectBytes: Uint8Array,
  dictionaryKind: PdfNativeResourceDictionaryKind,
  label: string,
  /** Read an owner that declares no `/XObject` at all as one naming no
   * pictures. Every resource category is optional (PDF 32000-1 7.8.3), and a
   * page of nothing but text -- which is most pages -- leaves the key out.
   * Callers that must name a picture the caller already selected keep the
   * default and refuse instead, so the refusal stays about that picture. */
  allowMissing = false,
): readonly PdfNativeReferenceDictionaryEntry[] {
  const resources = nativeResourceDictionarySource(
    ownerObjectBytes,
    dictionaryKind,
    label,
  );
  const xObjects = exactDirectDictionarySource(
    resources,
    "XObject",
    `${label} XObject resources`,
    allowMissing,
  );
  if (!xObjects) {
    if (allowMissing) return Object.freeze([]);
    fail("unsupported-document", `${label} XObject resources are missing.`);
  }
  // A page whose last picture was deleted keeps an empty `/XObject <<>>`;
  // that is a page with no images, not a malformed resource dictionary.
  return parseReferenceDictionarySource(
    xObjects,
    `${label} XObject resources`,
    "strict",
    true,
  );
}

/** The trailer's /Root is the Catalog whatever its dictionary says; a
 * producer may leave the /Type out (readers take the Root as the Catalog),
 * and one it writes has to say Catalog. Mirrors replay's catalog check. */
function isCatalogObject(objectSource: string): boolean {
  return (
    !oneTopLevelDictionaryEntry(objectSource, "Type", "PDF Catalog Type", true) ||
    exactObjectType(objectSource) === "Catalog"
  );
}

function exactObjectType(objectSource: string): "Catalog" | "Pages" | "Page" {
  const type = oneTopLevelName(
    objectSource,
    "Type",
    "PDF page-tree object Type",
  );
  if (type !== "Catalog" && type !== "Pages" && type !== "Page") {
    fail("unsupported-document", "PDF page-tree object Type is unsupported.");
  }
  return type;
}

const PAGE_BOX_NUMBER = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/;

/**
 * One page box (`MediaBox`, `CropBox`) declared on a page-tree node, as the
 * normalized `[x0, y0, x1, y1]`. Mirrors the server's `_top_level_page_box`
 * reading (a reference to a separate numeric array object, as TeX engines
 * write it, is followed), but a box the reader cannot parse is null rather
 * than a refusal: the vector scanners then treat every clip as real.
 */
function topLevelPageBox(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
  objectSource: string,
  key: "MediaBox" | "CropBox",
  readContext: NativeObjectReadContext,
): readonly [number, number, number, number] | null {
  let entry: TopLevelDictionaryEntry | null;
  try {
    entry = oneTopLevelDictionaryEntry(objectSource, key, `PDF page ${key}`, true);
  } catch (error) {
    if (error instanceof PdfNativeTransactionError) return null;
    throw error;
  }
  if (!entry) return null;
  let text = objectSource.slice(entry.valueStart, entry.valueEnd).trim();
  const reference = /^(\d+)\s+(\d+)\s+R$/.exec(text);
  if (reference) {
    try {
      const object = parseIndirectObject(
        sourceBytes,
        parsed,
        Number(reference[1]),
        Number(reference[2]),
        readContext,
      );
      if (object.dictionaryStart !== null) return null;
      text = asLatin1(object.bytes).trim();
    } catch (error) {
      if (error instanceof PdfNativeTransactionError) return null;
      throw error;
    }
  }
  if (!text.startsWith("[") || !text.endsWith("]")) return null;
  const values = text
    .slice(1, -1)
    .trim()
    .split(/\s+/)
    .filter((token) => token.length > 0);
  if (values.length !== 4 || values.some((token) => !PAGE_BOX_NUMBER.test(token))) {
    return null;
  }
  const numbers = values.map(Number);
  if (
    numbers.some((value) => !Number.isFinite(value) || Math.abs(value) > 1_000_000) ||
    numbers[2] <= numbers[0] ||
    numbers[3] <= numbers[1]
  ) {
    return null;
  }
  return Object.freeze([numbers[0], numbers[1], numbers[2], numbers[3]] as const);
}

function pageTreeKids(
  objectSource: string,
): readonly Readonly<{ objectNumber: number; generation: number }>[] {
  const entry = oneTopLevelDictionaryEntry(
    objectSource,
    "Kids",
    "PDF page-tree Kids",
  );
  if (!entry || objectSource[entry.valueStart] !== "[") {
    fail("unsupported-document", "PDF page-tree Kids are malformed.");
  }
  return parseTopLevelReferenceArray(
    objectSource,
    entry,
    "PDF page-tree Kids",
    1,
    4_096,
  );
}

function pageTreeCount(objectSource: string): number {
  const count = oneTopLevelInteger(
    objectSource,
    "Count",
    "PDF page-tree Count",
  );
  if (!Number.isSafeInteger(count) || count < 1 || count > 100_000) {
    fail("unsupported-document", "PDF page-tree Count is unsupported.");
  }
  return count;
}

interface TopLevelDictionaryEntry {
  readonly key: string;
  readonly keyStart: number;
  readonly valueStart: number;
  readonly valueEnd: number;
}

export interface PdfNativeTopLevelDictionaryNameEntry {
  readonly key: string;
  readonly directName: string | null;
}

function pdfTokenEnd(source: string, start: number, limit: number): number {
  let cursor = start;
  while (
    cursor < limit &&
    !isPdfWhitespace(source[cursor]) &&
    !"()<>[]{}/%".includes(source[cursor])
  ) {
    cursor += 1;
  }
  if (cursor === start) {
    fail("unsupported-document", "PDF dictionary token is malformed.");
  }
  return cursor;
}

function pdfNameEnd(source: string, start: number, limit: number): number {
  if (source[start] !== "/") {
    fail("unsupported-document", "PDF dictionary name is malformed.");
  }
  let cursor = start + 1;
  while (
    cursor < limit &&
    !isPdfWhitespace(source[cursor]) &&
    !"()<>[]{}/%".includes(source[cursor])
  ) {
    cursor += 1;
  }
  if (cursor === start + 1) {
    fail("unsupported-document", "PDF dictionary name is empty.");
  }
  return cursor;
}

function decodePdfName(source: string, start: number, end: number): string {
  if (end - start - 1 > MAX_PDF_NAME_BYTES) {
    fail("unsupported-document", "PDF dictionary name is too long.");
  }
  const decoded: string[] = [];
  for (let cursor = start + 1; cursor < end; cursor += 1) {
    if (source[cursor] !== "#") {
      decoded.push(source[cursor]);
      continue;
    }
    const escaped = source.slice(cursor + 1, cursor + 3);
    if (!/^[0-9A-Fa-f]{2}$/.test(escaped)) {
      fail("unsupported-document", "PDF dictionary name escape is malformed.");
    }
    decoded.push(String.fromCharCode(Number.parseInt(escaped, 16)));
    cursor += 2;
  }
  const value = decoded.join("");
  if (value.length > MAX_PDF_NAME_BYTES) {
    fail("unsupported-document", "PDF dictionary name is too long.");
  }
  return value;
}

function pdfLiteralEnd(source: string, start: number, limit: number): number {
  let depth = 0;
  let escaped = false;
  for (let cursor = start; cursor < limit; cursor += 1) {
    const character = source[cursor];
    if (escaped) {
      escaped = false;
      continue;
    }
    if (character === "\\") {
      escaped = true;
      continue;
    }
    if (character === "(") depth += 1;
    if (character === ")") {
      depth -= 1;
      if (depth === 0) return cursor + 1;
    }
  }
  fail("unsupported-document", "PDF dictionary literal is unterminated.");
}

function pdfHexEnd(source: string, start: number, limit: number): number {
  const end = source.indexOf(">", start + 1);
  if (end < 0 || end >= limit) {
    fail("unsupported-document", "PDF dictionary hex string is unterminated.");
  }
  return end + 1;
}

function pdfArrayEnd(source: string, start: number, limit: number): number {
  let depth = 0;
  for (let cursor = start; cursor < limit; cursor += 1) {
    const character = source[cursor];
    const next = source[cursor + 1];
    if (character === "%") {
      while (
        cursor < limit &&
        source[cursor] !== "\n" &&
        source[cursor] !== "\r"
      ) {
        cursor += 1;
      }
      continue;
    }
    if (character === "(") {
      cursor = pdfLiteralEnd(source, cursor, limit) - 1;
      continue;
    }
    if (character === "<" && next === "<") {
      cursor = dictionaryEnd(source, cursor) - 1;
      continue;
    }
    if (character === "<") {
      cursor = pdfHexEnd(source, cursor, limit) - 1;
      continue;
    }
    if (character === "[") depth += 1;
    if (character === "]") {
      depth -= 1;
      if (depth === 0) return cursor + 1;
    }
  }
  fail("unsupported-document", "PDF dictionary array is unterminated.");
}

function pdfDictionaryValueEnd(
  source: string,
  start: number,
  limit: number,
): number {
  if (source.slice(start, start + 2) === "<<") {
    return dictionaryEnd(source, start);
  }
  if (source[start] === "[") return pdfArrayEnd(source, start, limit);
  if (source[start] === "(") return pdfLiteralEnd(source, start, limit);
  if (source[start] === "<") return pdfHexEnd(source, start, limit);
  if (source[start] === "/") return pdfNameEnd(source, start, limit);
  const firstEnd = pdfTokenEnd(source, start, limit);
  const first = source.slice(start, firstEnd);
  if (!/^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(first)) return firstEnd;
  const secondStart = skipWhitespace(source, firstEnd);
  if (secondStart >= limit || !/[+\-.0-9]/.test(source[secondStart])) {
    return firstEnd;
  }
  const secondEnd = pdfTokenEnd(source, secondStart, limit);
  if (!/^[+-]?\d+$/.test(source.slice(secondStart, secondEnd))) return firstEnd;
  const thirdStart = skipWhitespace(source, secondEnd);
  if (source[thirdStart] !== "R") return firstEnd;
  const thirdEnd = thirdStart + 1;
  if (
    thirdEnd < limit &&
    !isPdfWhitespace(source[thirdEnd]) &&
    !"()<>[]{}/%".includes(source[thirdEnd])
  ) {
    return firstEnd;
  }
  return thirdEnd;
}

/** An inline image's own dictionary, read the way a reader reads it: bare
 * `/Key value` pairs between `BI` and `ID`, each value kept as its source
 * text. Anything the lexer refuses gives back no entries at all, so a caller
 * falls back to reading where the data ends from the stream itself. */
function inlineImageDictionaryEntries(
  dictionary: string,
): ReadonlyMap<string, string> {
  const entries = new Map<string, string>();
  try {
    let cursor = 0;
    while (cursor < dictionary.length) {
      cursor = skipWhitespace(dictionary, cursor);
      if (cursor >= dictionary.length) break;
      if (dictionary[cursor] !== "/") return new Map();
      const keyEnd = pdfNameEnd(dictionary, cursor, dictionary.length);
      const key = decodePdfName(dictionary, cursor, keyEnd);
      const valueStart = skipWhitespace(dictionary, keyEnd);
      if (valueStart >= dictionary.length) return new Map();
      const valueEnd = pdfDictionaryValueEnd(
        dictionary,
        valueStart,
        dictionary.length,
      );
      if (entries.size >= MAX_TOP_LEVEL_DICTIONARY_ENTRIES) return new Map();
      if (!entries.has(key)) {
        entries.set(key, dictionary.slice(valueStart, valueEnd));
      }
      cursor = valueEnd;
    }
  } catch (error) {
    if (!(error instanceof PdfNativeTransactionError)) throw error;
    return new Map();
  }
  return entries;
}

/** How many samples one pixel of an inline image's colour space carries. */
const INLINE_IMAGE_COLOR_SPACE_COMPONENTS = new Map<string, number>([
  ["G", 1],
  ["DeviceGray", 1],
  ["CalGray", 1],
  ["RGB", 3],
  ["DeviceRGB", 3],
  ["CalRGB", 3],
  ["CMYK", 4],
  ["DeviceCMYK", 4],
]);

/** The byte count an inline image's own dictionary states, or null when it
 * states none this engine can count: `/L` (`/Length`) outright, else the exact
 * sample count of an image no filter compresses. */
function inlineImageDeclaredDataLength(dictionary: string): number | null {
  const entries = inlineImageDictionaryEntries(dictionary);
  const read = (abbreviation: string, full: string): string | undefined =>
    entries.get(abbreviation) ?? entries.get(full);
  const declared = read("L", "Length");
  if (declared !== undefined) {
    return /^\d{1,10}$/.test(declared) ? Number(declared) : null;
  }
  const filter = read("F", "Filter");
  if (filter !== undefined && !/^\[\s*\]$/.test(filter)) return null;
  const width = read("W", "Width");
  const height = read("H", "Height");
  if (
    width === undefined ||
    height === undefined ||
    !/^\d{1,7}$/.test(width) ||
    !/^\d{1,7}$/.test(height)
  ) {
    return null;
  }
  const stencil = read("IM", "ImageMask") === "true";
  const bits = stencil ? "1" : (read("BPC", "BitsPerComponent") ?? "8");
  if (!/^(?:1|2|4|8|16)$/.test(bits)) return null;
  const space = read("CS", "ColorSpace");
  const components = stencil
    ? 1
    : space === undefined
      ? null
      : space.startsWith("/")
        ? (INLINE_IMAGE_COLOR_SPACE_COMPONENTS.get(space.slice(1)) ?? null)
        : /^\[\s*\/(?:I|Indexed)[\s/[(<]/.test(space)
          ? 1
          : null;
  if (components === null) return null;
  const rowBytes = Math.ceil(
    (Number(width) * components * Number(bits)) / 8,
  );
  return rowBytes * Number(height);
}

/** Where an inline image's binary data ends: the first `EI` that stands on
 * its own between whitespace and is followed by readable content. The data
 * starts one whitespace byte after `ID` and may contain anything, `EI`
 * included, so where the image's own dictionary states how many bytes it
 * wrote -- `/L`, or the sample count of an unfiltered image -- that count is
 * believed first and only the `EI` it lands on is taken; the free scan is the
 * rule pdf.js falls back to when nothing states the count. Replay applies the
 * same rule. */
export function mraiNativeInlineImageDataEnd(
  source: string,
  dataStart: number,
  dictionary?: string,
): number | null {
  const declared =
    dictionary === undefined ? null : inlineImageDeclaredDataLength(dictionary);
  if (declared !== null && dataStart + declared <= source.length) {
    // Only whitespace separates the stated data from its `EI`; a `%` here is
    // content, never a comment.
    let at = dataStart + declared;
    while (at < source.length && isPdfWhitespace(source[at])) at += 1;
    const after = source[at + 2];
    if (
      source.slice(at, at + 2) === "EI" &&
      (after === undefined ||
        isPdfWhitespace(after) ||
        "()<>[]{}/%".includes(after))
    ) {
      return at + 2;
    }
  }
  let cursor = dataStart;
  while (cursor < source.length) {
    const at = source.indexOf("EI", cursor);
    if (at < 0) return null;
    const after = source[at + 2];
    if (
      at > dataStart &&
      isPdfWhitespace(source[at - 1]) &&
      (after === undefined || isPdfWhitespace(after)) &&
      [...source.slice(at + 2, at + 18)].every((character) => {
        const code = character.charCodeAt(0);
        return (
          code === 0x0a ||
          code === 0x0d ||
          code === 0x09 ||
          code === 0x0c ||
          code === 0x00 ||
          (code >= 0x20 && code <= 0x7e)
        );
      })
    ) {
      return at + 2;
    }
    cursor = at + 2;
  }
  return null;
}

/** A dictionary's entries by its text: the planner reads a page's font
 * dictionaries once per item, font and route, and the text is the key. A
 * dictionary that fails to parse fails again. */
const keptDictionaryEntries = new Map<string, readonly TopLevelDictionaryEntry[]>();
const keptNameEntries = new Map<string, readonly PdfNativeTopLevelDictionaryNameEntry[]>();
const MAX_KEPT_DICTIONARIES = 512;
function keptByText<T>(store: Map<string, T>, text: string, read: () => T): T {
  const known = store.get(text);
  if (known !== undefined) return known;
  const value = read();
  if (store.size >= MAX_KEPT_DICTIONARIES) store.delete(store.keys().next().value!);
  store.set(text, value);
  return value;
}
function topLevelDictionaryEntries(
  objectSource: string,
): readonly TopLevelDictionaryEntry[] {
  return keptByText(keptDictionaryEntries, objectSource, () => Object.freeze(topLevelDictionaryEntriesUncached(objectSource)));
}
function topLevelDictionaryEntriesUncached(
  objectSource: string,
): readonly TopLevelDictionaryEntry[] {
  const header =
    /^\d+[\x00\x09\x0a\x0c\x0d\x20]+\d+[\x00\x09\x0a\x0c\x0d\x20]+obj\b/.exec(
      objectSource,
    );
  const dictionaryStart = objectSource.startsWith("<<")
    ? 0
    : skipWhitespace(objectSource, header?.[0].length ?? 0);
  if (objectSource.slice(dictionaryStart, dictionaryStart + 2) !== "<<") {
    fail("unsupported-document", "PDF object dictionary is missing.");
  }
  const dictionaryLimit = dictionaryEnd(objectSource, dictionaryStart) - 2;
  const entries: TopLevelDictionaryEntry[] = [];
  let cursor = dictionaryStart + 2;
  while (true) {
    cursor = skipWhitespace(objectSource, cursor);
    if (cursor >= dictionaryLimit) break;
    if (objectSource[cursor] !== "/") {
      fail("unsupported-document", "PDF dictionary entry is malformed.");
    }
    const keyEnd = pdfNameEnd(objectSource, cursor, dictionaryLimit);
    const key = decodePdfName(objectSource, cursor, keyEnd);
    const valueStart = skipWhitespace(objectSource, keyEnd);
    if (valueStart >= dictionaryLimit) {
      fail("unsupported-document", `PDF dictionary ${key} value is missing.`);
    }
    const valueEnd = pdfDictionaryValueEnd(
      objectSource,
      valueStart,
      dictionaryLimit,
    );
    if (entries.length >= MAX_TOP_LEVEL_DICTIONARY_ENTRIES) {
      fail("unsupported-document", "PDF dictionary has too many entries.");
    }
    entries.push(Object.freeze({ key, keyStart: cursor, valueStart, valueEnd }));
    cursor = valueEnd;
  }
  return Object.freeze(entries);
}

/**
 * Preserve the ordered raw dictionary entry sequence while decoding PDF name
 * escapes. Callers that authenticate a small direct-name envelope must see
 * duplicates before any map-like PDF object representation can overwrite
 * them.
 */
/** Read one top-level `/Key true|false`; null when absent or indirect. */
export function parseMraiNativeTopLevelDictionaryBoolean(
  objectBytes: Uint8Array,
  key: string,
): boolean | null {
  const objectSource = asLatin1(objectBytes);
  const entries = topLevelDictionaryEntries(objectSource).filter(
    (entry) => entry.key === key,
  );
  if (entries.length !== 1) return null;
  const raw = objectSource
    .slice(entries[0].valueStart, entries[0].valueEnd)
    .trim();
  return raw === "true" ? true : raw === "false" ? false : null;
}

export function parseMraiNativeTopLevelDictionaryNameEntries(
  objectBytes: Uint8Array,
): readonly PdfNativeTopLevelDictionaryNameEntry[] {
  const objectSource = asLatin1(objectBytes);
  return keptByText(keptNameEntries, objectSource, () => parseTopLevelDictionaryNameEntriesUncached(objectSource));
}
function parseTopLevelDictionaryNameEntriesUncached(
  objectSource: string,
): readonly PdfNativeTopLevelDictionaryNameEntry[] {
  return Object.freeze(
    topLevelDictionaryEntries(objectSource).map((entry) => {
      const directName = (() => {
        if (objectSource[entry.valueStart] !== "/") return null;
        const end = pdfNameEnd(objectSource, entry.valueStart, entry.valueEnd);
        return end === entry.valueEnd
          ? decodePdfName(objectSource, entry.valueStart, end)
          : null;
      })();
      return Object.freeze({
        key: entry.key,
        directName,
      });
    }),
  );
}

function oneTopLevelDictionaryEntry(
  objectSource: string,
  key: string,
  label: string,
  allowMissing = false,
): TopLevelDictionaryEntry | null {
  const matches = topLevelDictionaryEntries(objectSource).filter(
    (entry) => entry.key === key,
  );
  if (allowMissing && matches.length === 0) return null;
  if (matches.length !== 1) {
    fail("unsupported-document", `${label} is missing or ambiguous.`);
  }
  return matches[0];
}

function skipWhitespaceWithin(
  source: string,
  start: number,
  limit: number,
): number {
  let cursor = start;
  while (cursor < limit) {
    if (isPdfWhitespace(source[cursor])) {
      cursor += 1;
      continue;
    }
    if (source[cursor] === "%") {
      while (
        cursor < limit &&
        source[cursor] !== "\n" &&
        source[cursor] !== "\r"
      ) {
        cursor += 1;
      }
      continue;
    }
    break;
  }
  return cursor;
}

function parseTopLevelReference(
  source: string,
  start: number,
  limit: number,
  label: string,
): Readonly<{
  reference: Readonly<{ objectNumber: number; generation: number }>;
  end: number;
}> {
  let cursor = skipWhitespaceWithin(source, start, limit);
  const objectEnd = pdfTokenEnd(source, cursor, limit);
  const objectToken = source.slice(cursor, objectEnd);
  if (!/^\d+$/.test(objectToken)) {
    fail("unsupported-document", `${label} object number is malformed.`);
  }
  cursor = skipWhitespaceWithin(source, objectEnd, limit);
  const generationEnd = pdfTokenEnd(source, cursor, limit);
  const generationToken = source.slice(cursor, generationEnd);
  if (!/^\d+$/.test(generationToken)) {
    fail("unsupported-document", `${label} generation is malformed.`);
  }
  cursor = skipWhitespaceWithin(source, generationEnd, limit);
  const operatorEnd = pdfTokenEnd(source, cursor, limit);
  if (source.slice(cursor, operatorEnd) !== "R") {
    fail("unsupported-document", `${label} is not one indirect reference.`);
  }
  const objectNumber = Number(objectToken);
  const generation = Number(generationToken);
  assertObjectReference(objectNumber, generation);
  return Object.freeze({
    reference: Object.freeze({ objectNumber, generation }),
    end: operatorEnd,
  });
}

function parseTopLevelReferenceArray(
  source: string,
  entry: TopLevelDictionaryEntry,
  label: string,
  minimum: number,
  maximum: number,
): readonly Readonly<{ objectNumber: number; generation: number }>[] {
  if (source[entry.valueStart] !== "[" || source[entry.valueEnd - 1] !== "]") {
    fail("unsupported-document", `${label} is not one reference array.`);
  }
  const limit = entry.valueEnd - 1;
  let cursor = entry.valueStart + 1;
  const references: Array<
    Readonly<{ objectNumber: number; generation: number }>
  > = [];
  while (true) {
    cursor = skipWhitespaceWithin(source, cursor, limit);
    if (cursor >= limit) break;
    const parsed = parseTopLevelReference(source, cursor, limit, label);
    references.push(parsed.reference);
    cursor = parsed.end;
  }
  if (references.length < minimum || references.length > maximum) {
    fail("unsupported-document", `${label} size is unsupported.`);
  }
  if (
    new Set(references.map((item) => `${item.objectNumber}:${item.generation}`))
      .size !== references.length
  ) {
    fail("unsupported-document", `${label} contains duplicate identities.`);
  }
  return Object.freeze(references);
}

function oneTopLevelName(
  objectSource: string,
  key: string,
  label: string,
): string {
  const entry = oneTopLevelDictionaryEntry(objectSource, key, label);
  if (!entry || objectSource[entry.valueStart] !== "/") {
    fail("unsupported-document", `${label} is not one name.`);
  }
  const end = pdfNameEnd(objectSource, entry.valueStart, entry.valueEnd);
  if (end !== entry.valueEnd) {
    fail("unsupported-document", `${label} name is malformed.`);
  }
  return decodePdfName(objectSource, entry.valueStart, end);
}

export { oneTopLevelName as readMraiNativeDictionaryName };
export { dictionaryEnd as readMraiNativeDictionaryEnd };
export { topLevelDictionaryEntries as readMraiNativeTopLevelDictionaryEntries };

function oneTopLevelInteger(
  objectSource: string,
  key: string,
  label: string,
): number {
  const entry = oneTopLevelDictionaryEntry(objectSource, key, label);
  if (!entry) fail("unsupported-document", `${label} is missing.`);
  const value = objectSource.slice(entry.valueStart, entry.valueEnd);
  if (!/^\d+$/.test(value)) {
    fail("unsupported-document", `${label} is not one integer.`);
  }
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed)) {
    fail("unsupported-document", `${label} is outside the safe range.`);
  }
  return parsed;
}

function oneTopLevelReference(
  objectSource: string,
  key: string,
  label: string,
  allowMissing = false,
): Readonly<{ objectNumber: number; generation: number }> | null {
  const entry = oneTopLevelDictionaryEntry(
    objectSource,
    key,
    label,
    allowMissing,
  );
  if (!entry) return null;
  const parsed = parseTopLevelReference(
    objectSource,
    entry.valueStart,
    entry.valueEnd,
    label,
  );
  if (parsed.end !== entry.valueEnd) {
    fail("unsupported-document", `${label} is malformed.`);
  }
  return parsed.reference;
}

interface PageResourceAuthority {
  readonly owner: Readonly<{ objectNumber: number; generation: number }>;
  readonly dictionaryKind: PdfNativeResourceDictionaryKind;
}

function ownPageResourceAuthority(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
  owner: Readonly<{ objectNumber: number; generation: number }>,
  objectSource: string,
): PageResourceAuthority | null {
  const matches = topLevelDictionaryEntries(objectSource).filter(
    (entry) => entry.key === "Resources",
  );
  if (matches.length === 0) return null;
  if (matches.length !== 1) {
    fail("unsupported-document", "PDF page Resources are ambiguous.");
  }
  const value = objectSource
    .slice(matches[0].valueStart, matches[0].valueEnd)
    .trim();
  if (value.startsWith("<<")) {
    return Object.freeze({
      owner: Object.freeze({ ...owner }),
      dictionaryKind: "nested-resources",
    });
  }
  const parsedReference = parseTopLevelReference(
    objectSource,
    matches[0].valueStart,
    matches[0].valueEnd,
    "PDF page Resources",
  );
  if (parsedReference.end !== matches[0].valueEnd) {
    fail("unsupported-document", "PDF page Resources are malformed.");
  }
  const resourceOwner = parsedReference.reference;
  const resourceObject = parseIndirectObject(
    sourceBytes,
    parsed,
    resourceOwner.objectNumber,
    resourceOwner.generation,
  );
  if (resourceObject.streamBytes !== null) {
    fail(
      "unsupported-document",
      "Stream-backed page resource dictionaries are preserve-only.",
    );
  }
  const resourceSource = asLatin1(resourceObject.bytes);
  const resourceEntries = topLevelDictionaryEntries(resourceSource);
  if (
    resourceEntries.some((entry) =>
      ["Type", "Parent", "Kids", "Contents", "Resources"].includes(entry.key),
    )
  ) {
    fail(
      "unsupported-document",
      "Indirect page resource dictionary ownership is ambiguous.",
    );
  }
  return Object.freeze({
    owner: resourceOwner,
    dictionaryKind: "resource-dictionary",
  });
}

type PageGraphOptions = Readonly<{
  validateContentStreams?: boolean;
  includeExtGStates?: boolean;
  /** Record a page whose content cannot be read instead of failing the
   * whole document; only callers that judge that page on its own may ask. */
  tolerateUnusablePages?: boolean;
}>;

/** One source's page graphs, read once per buffer and set of options: a
 * plan walks every page of a document to count a shared form's instances,
 * and one edit plans several times (the native route, the font-reuse
 * check, the writer's proof), each of which asked for the same walk; a
 * 1200-page book took four minutes per walk. Sources are never written in
 * place, and the inspection is frozen. */
const pageGraphInspections = new WeakMap<
  Uint8Array,
  Map<string, Promise<readonly PdfNativePageContentGraphInspection[]>>
>();

async function inspectMraiNativeDocumentPageGraphs(
  sourceBytes: Uint8Array,
  options: PageGraphOptions = {},
): Promise<readonly PdfNativePageContentGraphInspection[]> {
  // One walk serves the strict and the tolerant callers alike: it records
  // a page it cannot read, and a strict caller refuses on the first one.
  const key = [
    options.validateContentStreams !== false,
    options.includeExtGStates === true,
  ].join("|");
  let byOptions = pageGraphInspections.get(sourceBytes);
  if (!byOptions) {
    byOptions = new Map();
    pageGraphInspections.set(sourceBytes, byOptions);
  }
  let pending = byOptions.get(key);
  if (!pending) {
    pending = inspectMraiNativeDocumentPageGraphsUncached(sourceBytes, {
      ...options,
      tolerateUnusablePages: true,
    });
    byOptions.set(key, pending);
    // A refused walk is not remembered: the next caller asks again.
    pending.catch(() => byOptions!.delete(key));
  }
  const pages = await pending;
  if (!options.tolerateUnusablePages) {
    const unusable = pages.find((page) => page.unusableContentError);
    if (unusable) throw unusable.unusableContentError;
  }
  return pages;
}

async function inspectMraiNativeDocumentPageGraphsUncached(
  sourceBytes: Uint8Array,
  options: PageGraphOptions,
): Promise<readonly PdfNativePageContentGraphInspection[]> {
  const validateContentStreams = options.validateContentStreams !== false;
  const parsed = parseSourceStructure(sourceBytes);
  const treeRead: NativeObjectReadContext = { source: asLatin1(sourceBytes), members: new Map() };
  const root = parseIndirectObject(
    sourceBytes,
    parsed,
    parsed.current.rootObjectNumber,
    parsed.current.rootGeneration,
    treeRead,
  );
  const rootSource = asLatin1(root.bytes);
  if (!isCatalogObject(rootSource)) {
    fail("unsupported-document", "PDF trailer Root is not a Catalog.");
  }
  const pagesRoot = oneTopLevelReference(
    rootSource,
    "Pages",
    "PDF Catalog Pages reference",
  );
  if (!pagesRoot) {
    fail("unsupported-document", "PDF Catalog Pages reference is missing.");
  }
  const leaves: Array<{
    ref: Readonly<{ objectNumber: number; generation: number }>;
    object: ParsedIndirectObject;
    resourceAuthority: PageResourceAuthority;
    visibleBox: readonly [number, number, number, number] | null;
  }> = [];
  const seen = new Set<string>();
  const visit = (
    ref: Readonly<{ objectNumber: number; generation: number }>,
    depth: number,
    parent: Readonly<{ objectNumber: number; generation: number }> | null,
    inheritedResources: PageResourceAuthority | null,
    inheritedMediaBox: readonly [number, number, number, number] | null,
    inheritedCropBox: readonly [number, number, number, number] | null,
  ): number => {
    if (depth > 32) {
      fail("unsupported-document", "PDF page tree is too deep.");
    }
    const key = `${ref.objectNumber}:${ref.generation}`;
    if (seen.has(key)) {
      fail("unsupported-document", "PDF page tree contains a cycle.");
    }
    if (seen.size >= 4_096) {
      fail("unsupported-document", "PDF page tree is too broad.");
    }
    seen.add(key);
    const object = parseIndirectObject(
      sourceBytes,
      parsed,
      ref.objectNumber,
      ref.generation,
      treeRead,
    );
    const objectSource = asLatin1(object.bytes);
    const type = exactObjectType(objectSource);
    const declaredParent = oneTopLevelReference(
      objectSource,
      "Parent",
      "PDF page-tree Parent reference",
      true,
    );
    if (parent === null) {
      if (declaredParent !== null) {
        fail(
          "unsupported-document",
          "PDF page-tree root cannot declare a Parent.",
        );
      }
    } else if (
      declaredParent === null ||
      declaredParent.objectNumber !== parent.objectNumber ||
      declaredParent.generation !== parent.generation
    ) {
      fail(
        "unsupported-document",
        "PDF page-tree Parent identity does not match its owner.",
      );
    }
    const resourceAuthority =
      ownPageResourceAuthority(sourceBytes, parsed, ref, objectSource) ??
      inheritedResources;
    const mediaBox =
      topLevelPageBox(sourceBytes, parsed, objectSource, "MediaBox", treeRead) ??
      inheritedMediaBox;
    const cropBox =
      topLevelPageBox(sourceBytes, parsed, objectSource, "CropBox", treeRead) ??
      inheritedCropBox;
    if (type === "Page") {
      // The server refuses a CropBox outside its MediaBox; here that page
      // simply has no visible box, which keeps every clip on it a real one.
      const visibleBox =
        mediaBox === null
          ? null
          : cropBox === null
            ? mediaBox
            : cropBox[0] < mediaBox[0] - 0.001 ||
                cropBox[1] < mediaBox[1] - 0.001 ||
                cropBox[2] > mediaBox[2] + 0.001 ||
                cropBox[3] > mediaBox[3] + 0.001
              ? null
              : cropBox;
      leaves.push({
        ref,
        object,
        resourceAuthority:
          resourceAuthority ??
          // A page that declares none and inherits none names none.
          Object.freeze({
            owner: Object.freeze({ ...ref }),
            dictionaryKind: "absent-resources" as const,
          }),
        visibleBox,
      });
      return 1;
    }
    if (type !== "Pages") {
      fail("unsupported-document", "PDF page tree contains a non-page node.");
    }
    const count = pageTreeKids(objectSource).reduce(
      (total, child) =>
        total +
        visit(child, depth + 1, ref, resourceAuthority, mediaBox, cropBox),
      0,
    );
    if (count !== pageTreeCount(objectSource)) {
      fail("unsupported-document", "PDF page-tree Count does not match Kids.");
    }
    return count;
  };
  visit(pagesRoot, 0, null, null, null, null);
  treeRead.source = "";
  treeRead.members.clear();
  const sourceSha256 = await sha256PdfBytes(sourceBytes);
  // Caller bytes are not owned here: never carry a decoded cache across an
  // await. This second synchronous traversal reads their current contents.
  const contentRead: NativeObjectReadContext = { source: asLatin1(sourceBytes), members: new Map() };
  const pages = Object.freeze(
    leaves.map((selected, index) => {
      const pageSource = asLatin1(selected.object.bytes);
      const contentStreams = resolvedPageContentReferences(
        sourceBytes,
        parsed,
        pageSource,
        contentRead,
      );
      let unusableContentStreams: string | undefined;
      let unusableContentError: PdfNativeTransactionError | undefined;
      if (validateContentStreams) {
        try {
          const decodedContentLengths: number[] = [];
          for (const contentStream of contentStreams) {
            const object = parseIndirectObject(
              sourceBytes,
              parsed,
              contentStream.objectNumber,
              contentStream.generation,
              contentRead,
            );
            assertTargetStream(object);
            decodedContentLengths.push(object.streamBytes.length);
            assertMraiNativePageDecodedContentBudget(decodedContentLengths);
          }
        } catch (error) {
          // One page whose content cannot be read must not close the whole
          // document: the caller decides which routes that page's state
          // affects.
          if (!options.tolerateUnusablePages || !(error instanceof PdfNativeTransactionError)) {
            throw error;
          }
          unusableContentStreams = `Page ${index + 1}: ${error.message}`;
          unusableContentError = error;
        }
      }
      return Object.freeze({
        ...(unusableContentStreams ? { unusableContentStreams, unusableContentError } : {}),
        sourceSha256,
        pageNumber: index + 1,
        page: Object.freeze({ ...selected.ref }),
        resourceAuthority: Object.freeze({
          owner: Object.freeze({ ...selected.resourceAuthority.owner }),
          dictionaryKind: selected.resourceAuthority.dictionaryKind,
        }),
        contentStreams,
        pageObjectBytes: Uint8Array.from(selected.object.bytes),
        visibleBox: selected.visibleBox,
      });
    }),
  );
  contentRead.source = "";
  contentRead.members.clear();
  if (!options.includeExtGStates) return pages;
  return Object.freeze(
    await Promise.all(
      pages.map(async (page) =>
        Object.freeze({
          ...page,
          extGStates: (
            await resolvePageExtGStateGraph(sourceBytes, parsed, page)
          ).inspection,
        }),
      ),
    ),
  );
}

/**
 * Resolve every visible page and its ordered content streams from the original
 * pinned xref/object graph. This is intentionally read-only and exists so a
 * mutable Form XObject can be proven to have exactly one document-wide use.
 */
export async function inspectMraiNativeDocumentPageContentGraphs(
  sourceBytes: Uint8Array,
  options?: Readonly<{ includeExtGStates?: boolean }>,
): Promise<readonly PdfNativePageContentGraphInspection[]> {
  return inspectMraiNativeDocumentPageGraphs(sourceBytes, {
    includeExtGStates: options?.includeExtGStates === true,
    tolerateUnusablePages: true,
  });
}

/**
 * Whether another page draws one of this page's content streams too. Split
 * Pages writes each part as a window on the parent's one stream, so an edit
 * written into it changes every part (user audit PDF 0071); such a page is
 * rewritten on a copy of its own, its streams joined as one new object the
 * way a page whose streams straddle a text object is.
 */
export async function mraiNativePageSharesContent(
  sourceBytes: Uint8Array,
  pageNumber: number,
): Promise<boolean> {
  const pages = await inspectMraiNativeDocumentPageContentGraphs(sourceBytes);
  const page = pages[pageNumber - 1];
  if (!page) return false;
  return page.contentStreams.some((stream) =>
    pages.some((other) => other !== page && other.contentStreams.some((candidate) =>
      candidate.objectNumber === stream.objectNumber && candidate.generation === stream.generation)));
}

/**
 * Resolve one visible page and its ordered content streams from the original
 * pinned xref/object graph. UI/pdf.js object hints never enter this path.
 */
export async function inspectMraiNativePageContentGraph(
  sourceBytes: Uint8Array,
  pageNumber: number,
): Promise<PdfNativePageContentGraphInspection> {
  if (!Number.isSafeInteger(pageNumber) || pageNumber < 1) {
    fail("invalid-request", "Invalid PDF page number.");
  }
  const pages = await inspectMraiNativeDocumentPageGraphs(sourceBytes);
  const selected = pages[pageNumber - 1];
  if (!selected) {
    fail("invalid-request", "Selected PDF page is out of range.");
  }
  return selected;
}

const PDF_PAGE_CONTENT_WHITESPACE = new Set([
  "\x00",
  "\t",
  "\n",
  "\f",
  "\r",
  " ",
]);
const PDF_PAGE_CONTENT_DELIMITERS = new Set([
  "(",
  ")",
  "<",
  ">",
  "[",
  "]",
  "{",
  "}",
  "/",
  "%",
]);

/**
 * A page drawn from several content streams is one drawing: a reader joins
 * the streams with whitespace and reads the result as one operator sequence.
 * Two neighbours may be read that way only when the join cannot weld two
 * lexical tokens into one, because every per-stream byte offset a writer
 * names is an offset into that one drawing.
 */
export function hasMraiNativeSafeContentBoundary(
  left: Uint8Array,
  right: Uint8Array,
): boolean {
  if (left.length === 0 || right.length === 0) return false;
  const leftToken = String.fromCharCode(left[left.length - 1]);
  const rightToken = String.fromCharCode(right[0]);
  return (
    PDF_PAGE_CONTENT_WHITESPACE.has(leftToken) ||
    PDF_PAGE_CONTENT_WHITESPACE.has(rightToken) ||
    PDF_PAGE_CONTENT_DELIMITERS.has(leftToken) ||
    PDF_PAGE_CONTENT_DELIMITERS.has(rightToken) ||
    endsWithGraphicsStateBracket(left)
  );
}

/**
 * Whether a stream's last token is `q` or `Q` standing alone. A stamping
 * tool wraps the page it stamps in a `q` stream before it and a `Q…` stream
 * after it, and the page's own stream often ends in a bare `Q` with no line
 * end (iText's header on a court filing: `…/Im1 Do\nQ` then `Q\nq\nBT…`).
 * A division between streams is a division between tokens (ISO 32000-1
 * 7.8.2), and the readers that join the streams with white space read it so.
 * pdf.js joins them with nothing and ends a token where the next character
 * would make a known operator into no operator; no operator but `q` and `Q`
 * themselves starts with either letter, so it too ends the token at the
 * division, whatever the next stream starts with. Only this reading is
 * widened: streams written out as one (consolidatedVirtualTarget) are joined
 * byte to byte and keep the strict rule.
 */
function endsWithGraphicsStateBracket(left: Uint8Array): boolean {
  if (left.length < 2) return false;
  const last = left[left.length - 1];
  if (last !== 0x51 && last !== 0x71) return false;
  const before = String.fromCharCode(left[left.length - 2]);
  return PDF_PAGE_CONTENT_WHITESPACE.has(before) || PDF_PAGE_CONTENT_DELIMITERS.has(before);
}

export interface PdfNativePageContentStreamInspection {
  readonly pageNumber: number;
  readonly page: Readonly<{ objectNumber: number; generation: number }>;
  readonly resourceAuthority: PdfNativePageContentGraphInspection["resourceAuthority"];
  /** The page's ordered /Contents, exactly as the reader concatenates them. */
  readonly contentStreams: readonly Readonly<{
    objectNumber: number;
    generation: number;
  }>[];
  readonly contentStream: Readonly<{ objectNumber: number; generation: number }>;
  /** Where the selected stream sits inside the page's own /Contents. */
  readonly streamIndex: number;
  /** Where this stream's first byte sits inside the whole page drawing. */
  readonly drawingOffset: number;
  readonly streamBytes: Uint8Array;
  readonly identity: PdfNativeObjectIdentity;
  readonly objectSha256: string;
  readonly streamSha256: string;
  readonly untouchedObjects: readonly PdfNativeObjectPrecondition[];
}

/**
 * Every content stream of one page, each named and pinned as its own writable
 * target.
 *
 * A page may be drawn from several streams; a reader joins them with
 * whitespace and reads the result as one operator sequence. Each entry pins
 * the page object and every other stream of the same page as untouched, so a
 * span expressed as an offset into one part of the drawing can never be
 * replayed against a drawing whose other parts moved. The join itself is
 * refused when it would weld two lexical tokens together, because per-stream
 * offsets would then not describe the drawing a reader sees.
 */
export interface PdfNativePageContentStreamOptions {
  readonly extraUntouched?: readonly Readonly<{
    objectNumber: number;
    generation: number;
    role?: PdfNativeObjectPrecondition["role"];
  }>[];
  /** The role the page object is pinned under; writers differ on it. */
  readonly pageRole?: PdfNativeObjectPrecondition["role"];
}

export async function inspectMraiNativePageContentStreams(
  sourceBytes: Uint8Array,
  pageNumber: number,
  options: PdfNativePageContentStreamOptions = {},
): Promise<readonly PdfNativePageContentStreamInspection[]> {
  const extraUntouched = options.extraUntouched ?? [];
  const pageRole = options.pageRole ?? "resource";
  const graph = await inspectMraiNativePageContentGraph(sourceBytes, pageNumber);
  const contentStreams = graph.contentStreams;
  if (contentStreams.length < 1) {
    fail("unsupported-document", "PDF page Contents is missing or ambiguous.");
  }
  const parsed = parsePdf(sourceBytes, false);
  const sourceSha256 = await sha256PdfBytes(sourceBytes);
  const frozenStreams = Object.freeze(
    contentStreams.map((reference) => Object.freeze({ ...reference })),
  );
  const inspections: PdfNativePageContentStreamInspection[] = [];
  let drawingOffset = 0;
  for (const [streamIndex, target] of contentStreams.entries()) {
    const untouched = [
      { ...graph.page, role: pageRole },
      ...contentStreams
        .filter((_, index) => index !== streamIndex)
        .map((sibling) => ({ ...sibling, role: "untouched-object" as const })),
      ...extraUntouched,
    ].filter(
      (item, index, items) =>
        !sameNativeObjectRef(item, target) &&
        items.findIndex((other) => sameNativeObjectRef(other, item)) === index,
    );
    const inspected = await inspectNativeStreamInReadSource(
      sourceBytes,
      parsed,
      sourceSha256,
      target,
      untouched,
    );
    inspections.push(
      Object.freeze({
        pageNumber: graph.pageNumber,
        page: Object.freeze({ ...graph.page }),
        resourceAuthority: graph.resourceAuthority,
        contentStreams: frozenStreams,
        contentStream: Object.freeze({ ...target }),
        streamIndex,
        drawingOffset,
        streamBytes: inspected.streamBytes,
        identity: inspected.identity,
        objectSha256: inspected.objectSha256,
        streamSha256: inspected.streamSha256,
        untouchedObjects: inspected.untouchedObjects,
      }),
    );
    // One separator byte stands for the whitespace a reader joins with.
    drawingOffset += inspected.streamBytes.length + 1;
  }
  for (let index = 1; index < inspections.length; index += 1) {
    if (
      !hasMraiNativeSafeContentBoundary(
        inspections[index - 1].streamBytes,
        inspections[index].streamBytes,
      )
    ) {
      fail(
        "unsupported-document",
        "PDF content-stream array splits a lexical token boundary.",
      );
    }
  }
  return Object.freeze(inspections);
}

/**
 * One page content stream, named and pinned, for a writer that edits a span
 * inside it. The caller names the stream that owns the operator it selected —
 * the text planner records that identity with every proof it hands on. A
 * caller that cannot name one still gets the historical contract of one exact
 * page content stream, refused with its own wording.
 */
export async function inspectMraiNativePageContentStream(
  sourceBytes: Uint8Array,
  pageNumber: number,
  contentStream: Readonly<{ objectNumber: number; generation: number }> | null,
  singleStreamRefusal: string,
  options: PdfNativePageContentStreamOptions = {},
): Promise<PdfNativePageContentStreamInspection> {
  const inspections = await inspectMraiNativePageContentStreams(
    sourceBytes,
    pageNumber,
    options,
  );
  if (!contentStream) {
    if (inspections.length !== 1) {
      fail("unsupported-document", singleStreamRefusal);
    }
    return inspections[0];
  }
  const matches = inspections.filter((inspection) =>
    sameNativeObjectRef(inspection.contentStream, contentStream),
  );
  if (matches.length === 0) {
    fail(
      "precondition-mismatch",
      "The named content stream does not draw this PDF page.",
    );
  }
  if (matches.length > 1) {
    fail(
      "unsupported-document",
      "PDF page Contents names one content stream more than once.",
    );
  }
  return matches[0];
}

/**
 * The page drawing: its content streams joined the way a reader joins them,
 * one whitespace byte apart, so every `drawingOffset` above indexes into it.
 */
export function joinNativePageDrawing(
  inspections: readonly Readonly<{
    drawingOffset: number;
    streamBytes: Uint8Array;
  }>[],
): Uint8Array {
  const last = inspections.at(-1);
  const drawing = new Uint8Array(
    last ? last.drawingOffset + last.streamBytes.length : 0,
  );
  for (const inspection of inspections) {
    drawing.set(inspection.streamBytes, inspection.drawingOffset);
    const separator = inspection.drawingOffset + inspection.streamBytes.length;
    if (separator < drawing.length) drawing[separator] = 0x0a;
  }
  return drawing;
}
export async function inspectMraiNativeDocumentPageExtGStates(
  sourceBytes: Uint8Array,
): Promise<readonly PdfNativePageExtGStateInspection[]> {
  const pages = await inspectMraiNativeDocumentPageGraphs(sourceBytes, {
    validateContentStreams: false,
    includeExtGStates: true,
  });
  return Object.freeze(
    pages.map(
      (page) =>
        page.extGStates ??
        fail("unsupported-document", "PDF ExtGState inspection is missing."),
    ),
  );
}

export async function inspectMraiNativePageExtGStates(
  sourceBytes: Uint8Array,
  pageNumber: number,
): Promise<PdfNativePageExtGStateInspection> {
  if (!Number.isSafeInteger(pageNumber) || pageNumber < 1) {
    fail("invalid-request", "Invalid PDF page number.");
  }
  const pages = await inspectMraiNativeDocumentPageExtGStates(sourceBytes);
  const page = pages[pageNumber - 1];
  if (!page) fail("invalid-request", "Selected PDF page is out of range.");
  return page;
}

function exactOpacity(value: number | null, label: string): number | null {
  if (
    value !== null &&
    (typeof value !== "number" ||
      !Number.isFinite(value) ||
      value < 0 ||
      value > 1)
  ) {
    fail("invalid-request", `${label} opacity is invalid.`);
  }
  return value;
}

function formatNativePdfNumber(value: number): string {
  const normalized = Object.is(value, -0) ? 0 : value;
  return normalized
    .toFixed(6)
    .replace(/\.(\d*?)0+$/, (_, kept: string) => (kept ? `.${kept}` : "")) // no lookbehind: Safari < 16.4
    .replace(/\.$/, "");
}

function exactRequestedBlendMode(
  value: PdfStandardBlendMode | null,
): PdfStandardBlendMode | null {
  if (value === null) return null;
  if (!isPdfStandardBlendMode(value)) {
    fail("invalid-request", "PDF blend mode is invalid.");
  }
  return value;
}

export async function planMraiNativePageExtGStateMutation(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly resourceName: string;
  readonly fillOpacity: number | null;
  readonly strokeOpacity: number | null;
  readonly blendMode: PdfStandardBlendMode | null;
}): Promise<PdfNativePageExtGStateMutationPlan> {
  if (!PDF_RESOURCE_NAME.test(input.resourceName)) {
    fail("invalid-request", "PDF ExtGState resource name is invalid.");
  }
  const fillOpacity = exactOpacity(input.fillOpacity, "Fill");
  const strokeOpacity = exactOpacity(input.strokeOpacity, "Stroke");
  const blendMode = exactRequestedBlendMode(input.blendMode);
  if (
    fillOpacity === null &&
    strokeOpacity === null &&
    blendMode === null
  ) {
    fail("invalid-request", "PDF ExtGState appearance change is empty.");
  }
  const inspection = await inspectMraiNativePageExtGStates(
    input.sourceBytes,
    input.pageNumber,
  );
  const existing = inspection.entries.filter(
    (entry) => entry.resourceName === input.resourceName,
  );
  if (existing.length > 1) {
    fail("unsupported-document", "PDF ExtGState resource name is ambiguous.");
  }
  const reuseExisting = existing.length === 1;
  if (
    reuseExisting &&
    (!existing[0].supportedAppearanceOnly ||
      existing[0].fillOpacity !== fillOpacity ||
      existing[0].strokeOpacity !== strokeOpacity ||
      existing[0].blendMode !== blendMode)
  ) {
    fail(
      "unsupported-document",
      "PDF ExtGState resource name collides with different state.",
    );
  }
  return Object.freeze({
    pageNumber: inspection.pageNumber,
    page: inspection.page,
    resourceOwner: inspection.resourceOwner,
    category: inspection.category,
    resourceName: input.resourceName,
    fillOpacity,
    strokeOpacity,
    blendMode,
    reuseExisting,
  });
}

export async function planMraiNativePageFontExtensionMutation(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly fontPackage:
    | PdfNativeFontExtensionPackage
    | PdfNativeFontAliasPackage;
  /** More served-face packages installed in the same write. */
  readonly additionalFontPackages?: readonly PdfNativeFontExtensionPackage[];
  /** New object numbers taken before the font: a consolidated page's
   * stream, or the clones of an isolated Form chain, come first. */
  readonly reservedObjects?: number;
  /** Install the font in this Form XObject's own resources instead of the
   * page's: the leaf of the form instance isolation the edit rides on. */
  readonly form?: PdfNativeIndirectObjectRef;
}): Promise<PdfNativePageFontExtensionMutationPlan> {
  const fontPackage = normalizePageFontPackage(input.fontPackage);
  const additionalFontPackages = normalizeAdditionalFontPackages(fontPackage, input.additionalFontPackages);
  const installed = pdfNativePageFontPackages({ fontPackage, additionalFontPackages });
  const pageNumber = exactPositiveInteger(
    input.pageNumber,
    "font extension page number",
  );
  const parsed = parsePdf(input.sourceBytes, false);
  const pages = await inspectMraiNativeDocumentPageGraphs(input.sourceBytes);
  const page = pages[pageNumber - 1];
  // The mutation adds a font to the page's resources; which of the page's
  // content streams the edit lives in is the text transaction's business.
  if (!page || page.pageNumber !== pageNumber || page.contentStreams.length < 1) {
    fail(
      "unsupported-document",
      "Native font extension requires one exact page content stream.",
    );
  }
  const formGraph = input.form
    ? await resolveFormFontGraph(input.sourceBytes, parsed, input.form)
    : null;
  const graph =
    formGraph ?? (await resolvePageFontGraph(input.sourceBytes, parsed, page));
  if (
    installed.some((item) =>
      graph.entries.some((entry) => entry.resourceName === item.resourceName) ||
      graph.names.has(item.resourceName))
  ) {
    fail("unsupported-document", "PDF Font resource name already exists.");
  }
  const latest = latestEntries(parsed);
  const firstNewObject = Math.max(
    parsed.current.size,
    Math.max(0, ...latest.keys()) + 1,
  );
  if (
    fontPackage.font.objectNumber !==
    firstNewObject + (input.reservedObjects ?? 0)
  ) {
    fail(
      "precondition-mismatch",
      "Native font package object allocation is stale.",
    );
  }
  if (input.form && formGraph) {
    return formFontExtensionPlan(page.pageNumber, input.form, formGraph, fontPackage, additionalFontPackages);
  }
  return Object.freeze({
    ...(additionalFontPackages ? { additionalFontPackages } : {}),
    pageNumber: page.pageNumber,
    page: Object.freeze({
      ...page.page,
      expectedObjectSha256: await sha256PdfBytes(graph.pageObject.bytes),
    }),
    resourceOwner: Object.freeze({
      ...page.resourceAuthority.owner,
      expectedObjectSha256: await sha256PdfBytes(
        graph.resourceOwnerObject.bytes,
      ),
      dictionaryKind: page.resourceAuthority.dictionaryKind,
    }),
    category: graph.category,
    fontPackage,
  });
}

/** The plan for a font installed in a Form's own resources: the form stands
 * where the page does, its resources dictionary where the page's authority
 * does. */
async function formFontExtensionPlan(
  pageNumber: number,
  form: PdfNativeIndirectObjectRef,
  graph: Awaited<ReturnType<typeof resolveFormFontGraph>>,
  fontPackage: PdfNativePageFontExtensionMutationPlan["fontPackage"],
  additionalFontPackages?: PdfNativePageFontExtensionMutationPlan["additionalFontPackages"],
): Promise<PdfNativePageFontExtensionMutationPlan> {
  const formExpected = Object.freeze({
    objectNumber: form.objectNumber,
    generation: form.generation,
    expectedObjectSha256: await sha256PdfBytes(graph.pageObject.bytes),
  });
  return Object.freeze({
    pageNumber,
    page: formExpected,
    resourceOwner: Object.freeze({
      objectNumber: graph.resourceOwnerObject.objectNumber,
      generation: graph.resourceOwnerObject.generation,
      expectedObjectSha256: await sha256PdfBytes(graph.resourceOwnerObject.bytes),
      dictionaryKind: graph.resourcesIndirect
        ? ("resource-dictionary" as const)
        : ("nested-resources" as const),
    }),
    category: graph.category,
    fontPackage,
    ...(additionalFontPackages ? { additionalFontPackages } : {}),
    form: formExpected,
  });
}

type PageFontReferenceAsk = Readonly<{
  resourceName: string;
  font: Readonly<{ objectNumber: number; generation: number }>;
}>;

/** A dictionary's entries as (key, value) texts, the white space round a
 * value aside: what two spellings of one dictionary are compared by. */
function dictionaryEntryTexts(dictionary: string): readonly (readonly [string, string])[] {
  return topLevelDictionaryEntries(dictionary).map((entry) => [
    entry.key,
    dictionary.slice(entry.valueStart, entry.valueEnd).replace(/^[\x00\t\n\f\r ]+|[\x00\t\n\f\r ]+$/gu, ""),
  ] as const);
}

/** `after` is `before` with the value of `key` alone another (or `key`
 * alone put in): every other entry stands as it stood, in its order. */
function dictionaryChangedInOneEntry(before: string, after: string, key: string): boolean {
  const was = dictionaryEntryTexts(before);
  const now = dictionaryEntryTexts(after);
  return was.filter(([name]) => name === key).length <= 1 &&
    now.filter(([name]) => name === key).length === 1 &&
    JSON.stringify(was.filter(([name]) => name !== key)) === JSON.stringify(now.filter(([name]) => name !== key));
}

/**
 * What a page font reference writes, derived from the source alone: the
 * page's effective Resources as a new object, their Font category holding
 * one entry more for each font and nothing else another, and the page
 * naming that object, nothing else of it another. Every rule of the
 * reference is held here, so the plan, the preparation and the reopened
 * candidate are all judged by it:
 *  - each font is an object the file holds now, a `/Type /Font` dictionary
 *    (no stream), none of them the page, its resource owner or its Font
 *    category, and none a font the page already names;
 *  - each name is one a resource may carry, used once, and no name the
 *    page's Font category writes (a repeated or badly spelled one too);
 *  - the new Resources object takes the first number the file has free;
 *  - the stream the write changes (`stream`) is one of this page's own, and
 *    no other page draws it: a page that drew it too would select a font
 *    it does not name.
 */
async function resolvePageFontReference(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
  pageNumber: number,
  asked: readonly PageFontReferenceAsk[],
  stream: Readonly<{ objectNumber: number; generation: number }>,
): Promise<Readonly<{
  plan: PdfNativePageFontReferencePlan;
  page: PdfNativePageContentGraphInspection;
  resourcesObject: Uint8Array;
  revisedPage: Uint8Array;
}>> {
  if (!Array.isArray(asked) || asked.length < 1 || asked.length > MAX_MRAI_NATIVE_PAGE_FONT_REFERENCES) {
    fail("invalid-request", "A page font reference names one to four fonts.");
  }
  const pages = await inspectMraiNativeDocumentPageGraphs(sourceBytes);
  const page = pages[exactPositiveInteger(pageNumber, "font reference page number") - 1];
  if (!page || page.pageNumber !== pageNumber || page.contentStreams.length < 1) {
    fail("unsupported-document", "A page font reference requires a page with content.");
  }
  if (page.resourceAuthority.dictionaryKind === "absent-resources") {
    fail("unsupported-document", "The page names no resources a font could be named in.");
  }
  if (
    page.contentStreams.filter((item) => sameNativeObjectRef(item, stream)).length !== 1 ||
    pages.some((other) => other.pageNumber !== page.pageNumber &&
      other.contentStreams.some((item) => sameNativeObjectRef(item, stream)))
  ) {
    fail("unsupported-document", "Page font reference stream ownership is shared or ambiguous.");
  }
  const graph = await resolvePageFontGraph(sourceBytes, parsed, page);
  const latest = latestEntries(parsed);
  const resources = Object.freeze({
    objectNumber: Math.max(parsed.current.size, Math.max(0, ...latest.keys()) + 1),
    generation: 0 as const,
  });
  if (resources.objectNumber > 9_999_999_999) {
    fail("unsupported-document", "Font resource object number is too large.");
  }
  const names = new Set<string>();
  const targets = new Set<string>();
  const fonts: PdfNativePageFontReferencePlan["fonts"][number][] = [];
  for (const item of asked) {
    assertObjectReference(item.font.objectNumber, item.font.generation);
    const key = `${item.font.objectNumber}:${item.font.generation}`;
    if (!PDF_RESOURCE_NAME.test(item.resourceName) || names.has(item.resourceName) || targets.has(key)) {
      fail("invalid-request", "A page font reference names each font once, by a resource name.");
    }
    names.add(item.resourceName);
    targets.add(key);
    if (graph.names.has(item.resourceName) || graph.entries.some((entry) => entry.resourceName === item.resourceName)) {
      fail("unsupported-document", "PDF Font resource name already exists.");
    }
    if (graph.entries.some((entry) => entry.inlineDictionary === undefined && sameNativeObjectRef(entry.target, item.font))) {
      fail("unsupported-document", "The page already names the referenced font.");
    }
    const standing = latest.get(item.font.objectNumber);
    if (!standing?.inUse || standing.generation !== item.font.generation) {
      fail("precondition-mismatch", "A referenced font is not an object of the file.");
    }
    if (
      sameNativeObjectRef(item.font, page.page) ||
      sameNativeObjectRef(item.font, page.resourceAuthority.owner) ||
      (graph.category.kind === "indirect" && sameNativeObjectRef(item.font, graph.category.target)) ||
      page.contentStreams.some((content) => sameNativeObjectRef(content, item.font))
    ) {
      fail("invalid-request", "A referenced font is an object the page is written by.");
    }
    const object = parseIndirectObject(sourceBytes, parsed, item.font.objectNumber, item.font.generation);
    const dictionary = object.streamBytes === null && object.streamStart === null
      ? topLevelDictionarySource(asLatin1(object.bytes), "Referenced font")
      : fail("unsupported-document", "A referenced font is not a font dictionary.");
    const type = oneTopLevelDictionaryEntry(dictionary, "Type", "Referenced font Type", true);
    if (!type || !/^\/Font$/u.test(dictionary.slice(type.valueStart, type.valueEnd).replace(/[\x00\t\n\f\r ]+/gu, ""))) {
      fail("unsupported-document", "A referenced font is not a font dictionary.");
    }
    fonts.push(Object.freeze({
      resourceName: item.resourceName,
      font: Object.freeze({
        objectNumber: item.font.objectNumber,
        generation: item.font.generation,
        expectedObjectSha256: await sha256PdfBytes(object.bytes),
      }),
    }));
  }
  let categoryDictionary = graph.categoryDictionary;
  for (const item of fonts) {
    categoryDictionary = rewriteDictionaryEntry(
      categoryDictionary,
      item.resourceName,
      `${item.font.objectNumber} ${item.font.generation} R`,
      "PDF Font resource name",
    );
  }
  const resourcesDictionary = rewriteDictionaryEntry(
    graph.resourcesDictionary,
    "Font",
    categoryDictionary,
    "PDF Font category",
  );
  // Read again as written: the category holds its entries as they stood and
  // these after them, and the Resources nothing else another.
  const writtenCategory = oneTopLevelDictionaryEntry(resourcesDictionary, "Font", "PDF Font category");
  if (
    !writtenCategory ||
    !dictionaryChangedInOneEntry(graph.resourcesDictionary, resourcesDictionary, "Font") ||
    JSON.stringify(dictionaryEntryTexts(resourcesDictionary.slice(writtenCategory.valueStart, writtenCategory.valueEnd))) !==
      JSON.stringify([
        ...(graph.categoryDictionary === "<<>>" ? [] : dictionaryEntryTexts(graph.categoryDictionary)),
        ...fonts.map((item) => [item.resourceName, `${item.font.objectNumber} ${item.font.generation} R`] as const),
      ])
  ) {
    fail("unsupported-document", "The page's resources cannot take a font entry and stay as they are otherwise.");
  }
  const pageSource = asLatin1(graph.pageObject.bytes);
  const revisedPageSource = rewriteDictionaryEntry(
    pageSource,
    "Resources",
    `${resources.objectNumber} 0 R`,
    "PDF page Resources",
  );
  const pageResources = oneTopLevelDictionaryEntry(revisedPageSource, "Resources", "PDF page Resources");
  if (
    !pageResources ||
    !dictionaryChangedInOneEntry(pageSource, revisedPageSource, "Resources") ||
    revisedPageSource.slice(pageResources.valueStart, pageResources.valueEnd).replace(/^[\x00\t\n\f\r ]+|[\x00\t\n\f\r ]+$/gu, "") !==
      `${resources.objectNumber} 0 R`
  ) {
    fail("unsupported-document", "The page cannot name new resources and stay as it is otherwise.");
  }
  const revisedPageValue = latin1Bytes(revisedPageSource);
  return Object.freeze({
    plan: Object.freeze({
      pageNumber: page.pageNumber,
      page: Object.freeze({
        objectNumber: page.page.objectNumber,
        generation: page.page.generation,
        expectedObjectSha256: await sha256PdfBytes(graph.pageObject.bytes),
      }),
      resourceOwner: Object.freeze({
        objectNumber: page.resourceAuthority.owner.objectNumber,
        generation: page.resourceAuthority.owner.generation,
        expectedObjectSha256: await sha256PdfBytes(graph.resourceOwnerObject.bytes),
        dictionaryKind: page.resourceAuthority.dictionaryKind,
      }),
      category: graph.category,
      resources,
      fonts: Object.freeze(fonts),
    }),
    page,
    resourcesObject: latin1Bytes(`${resources.objectNumber} 0 obj\n${resourcesDictionary}\nendobj`),
    revisedPage:
      graph.pageObject.storage === "direct"
        ? revisedPageValue
        : concatBytes([
            ascii(`${graph.pageObject.objectNumber} ${graph.pageObject.generation} obj\n`),
            revisedPageValue,
            ascii("\nendobj"),
          ]),
  });
}

/** One spelling of a reference plan, for comparing a request's with the one
 * the source gives now. */
function pageFontReferenceSummary(plan: PdfNativePageFontReferencePlan): string {
  return JSON.stringify({
    pageNumber: plan.pageNumber,
    page: [plan.page.objectNumber, plan.page.generation, plan.page.expectedObjectSha256],
    resourceOwner: [plan.resourceOwner.objectNumber, plan.resourceOwner.generation,
      plan.resourceOwner.expectedObjectSha256, plan.resourceOwner.dictionaryKind],
    category: [plan.category.kind, plan.category.expectedDictionarySha256,
      ...(plan.category.kind === "indirect"
        ? [plan.category.target.objectNumber, plan.category.target.generation, plan.category.target.expectedObjectSha256]
        : [])],
    resources: [plan.resources.objectNumber, plan.resources.generation],
    fonts: plan.fonts.map((item) => [item.resourceName, item.font.objectNumber, item.font.generation,
      item.font.expectedObjectSha256]),
  });
}

/**
 * The plan that names fonts of the file in one page's resources
 * (PdfNativePageFontReferencePlan). Each takes a name of its own the page's
 * Font category does not write: `MRF` and its object number, with a count
 * after it where that name is taken.
 */
export async function planMraiNativePageFontReference(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly fonts: readonly Readonly<{ objectNumber: number; generation: number }>[];
  /** The content stream the write changes: one of the page's own that no
   * other page draws, or the reference is refused. */
  readonly stream: Readonly<{ objectNumber: number; generation: number }>;
}): Promise<PdfNativePageFontReferencePlan> {
  const parsed = parsePdf(input.sourceBytes, false);
  const pages = await inspectMraiNativeDocumentPageGraphs(input.sourceBytes);
  const page = pages[exactPositiveInteger(input.pageNumber, "font reference page number") - 1];
  if (!page || page.pageNumber !== input.pageNumber) {
    fail("invalid-request", "Selected PDF page is out of range.");
  }
  if (page.resourceAuthority.dictionaryKind === "absent-resources") {
    fail("unsupported-document", "The page names no resources a font could be named in.");
  }
  const graph = await resolvePageFontGraph(input.sourceBytes, parsed, page);
  const taken = new Set([...graph.names, ...graph.entries.map((entry) => entry.resourceName)]);
  const asked = input.fonts.map((font) => {
    const stem = `MRF${font.objectNumber}${font.generation ? `g${font.generation}` : ""}`;
    let resourceName = stem;
    for (let count = 2; taken.has(resourceName); count += 1) resourceName = `${stem}_${count}`;
    taken.add(resourceName);
    return Object.freeze({ resourceName, font });
  });
  return (await resolvePageFontReference(input.sourceBytes, parsed, input.pageNumber, asked, input.stream)).plan;
}

const documentFontReferences = new WeakMap<Uint8Array, Promise<readonly Readonly<{
  font: Readonly<{ objectNumber: number; generation: number }>;
  pages: readonly number[];
}>[]>>();

/**
 * Every font object the file's pages name by reference in their Font
 * resources, in the order of their object numbers, each with the pages that
 * name it: where a face one page lacks is looked for in the rest of the
 * file. A page whose resources cannot be read names none here. Read-only,
 * once per source buffer.
 */
export function readMraiNativeDocumentFontReferences(sourceBytes: Uint8Array): Promise<readonly Readonly<{
  font: Readonly<{ objectNumber: number; generation: number }>;
  pages: readonly number[];
}>[]> {
  const known = documentFontReferences.get(sourceBytes);
  if (known) return known;
  const read = (async () => {
    const parsed = parsePdf(sourceBytes, false);
    const pages = await inspectMraiNativeDocumentPageGraphs(sourceBytes);
    const found = new Map<string, { font: Readonly<{ objectNumber: number; generation: number }>; pages: number[] }>();
    // (Pages that share or inherit one Resources object are read once.)
    const byOwner = new Map<string, readonly PdfNativeReferenceDictionaryEntry[]>();
    for (const page of pages) {
      if (page.resourceAuthority.dictionaryKind === "absent-resources") continue;
      const owner = `${page.resourceAuthority.owner.objectNumber}:${page.resourceAuthority.owner.generation}:${page.resourceAuthority.dictionaryKind}`;
      let entries = byOwner.get(owner);
      if (!entries) {
        try {
          entries = (await resolvePageFontGraph(sourceBytes, parsed, page)).entries;
        } catch (error) {
          if (!(error instanceof PdfNativeTransactionError)) throw error;
          entries = Object.freeze([]);
        }
        byOwner.set(owner, entries);
      }
      for (const entry of entries) {
        if (entry.inlineDictionary !== undefined) continue;
        const key = `${entry.target.objectNumber}:${entry.target.generation}`;
        const held = found.get(key) ?? { font: Object.freeze({ ...entry.target }), pages: [] };
        if (held.pages.at(-1) !== page.pageNumber) held.pages.push(page.pageNumber);
        found.set(key, held);
      }
    }
    return Object.freeze([...found.values()]
      .sort((left, right) => left.font.objectNumber - right.font.objectNumber || left.font.generation - right.font.generation)
      .map((item) => Object.freeze({ font: item.font, pages: Object.freeze(item.pages) })));
  })();
  documentFontReferences.set(sourceBytes, read);
  read.catch(() => { if (documentFontReferences.get(sourceBytes) === read) documentFontReferences.delete(sourceBytes); });
  return read;
}

export async function inspectMraiNativeReusablePageFontExtension(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly replacementText: string;
  readonly fontBytes: Uint8Array | ArrayBuffer;
}): Promise<PdfNativeReusablePageFontExtension | null> {
  const sourceSha256 = await sha256PdfBytes(input.sourceBytes);
  const pageNumber = exactPositiveInteger(
    input.pageNumber,
    "reusable font page number",
  );
  const parsed = parsePdf(input.sourceBytes, false);
  const pages = await inspectMraiNativeDocumentPageGraphs(input.sourceBytes);
  const page = pages[pageNumber - 1];
  if (!page || page.pageNumber !== pageNumber) {
    fail("unsupported-document", "Reusable font page is missing.");
  }
  const graph = await resolvePageFontGraph(input.sourceBytes, parsed, page);
  const candidates = graph.entries
    .filter(
      (entry) =>
        entry.resourceName.startsWith("MRAI_NF_") &&
        entry.target.generation === 0,
    )
    .sort(
      (left, right) =>
        left.target.objectNumber - right.target.objectNumber ||
        left.resourceName.localeCompare(right.resourceName),
    );
  for (const entry of candidates) {
    try {
      const fontObject = parseIndirectObject(
        input.sourceBytes,
        parsed,
        entry.target.objectNumber,
        entry.target.generation,
      );
      const fontDictionary = topLevelDictionarySource(
        asLatin1(fontObject.bytes),
        "reusable Type0 font",
      );
      const toUnicodeEntry = oneTopLevelDictionaryEntry(
        fontDictionary,
        "ToUnicode",
        "reusable Type0 ToUnicode",
      );
      if (!toUnicodeEntry) continue;
      const toUnicodeRef = parseTopLevelReference(
        fontDictionary,
        toUnicodeEntry.valueStart,
        toUnicodeEntry.valueEnd,
        "reusable Type0 ToUnicode",
      ).reference;
      const descendantEntry = oneTopLevelDictionaryEntry(
        fontDictionary,
        "DescendantFonts",
        "reusable Type0 descendant",
      );
      if (!descendantEntry) continue;
      const descendants = parseTopLevelReferenceArray(
        fontDictionary,
        descendantEntry,
        "reusable Type0 descendant",
        1,
        1,
      );
      const descendantRef = descendants[0];
      const descendantObject = parseIndirectObject(
        input.sourceBytes,
        parsed,
        descendantRef.objectNumber,
        descendantRef.generation,
      );
      const descendantDictionary = topLevelDictionarySource(
        asLatin1(descendantObject.bytes),
        "reusable CIDFont",
      );
      const descriptorRef = oneTopLevelReference(
        descendantDictionary,
        "FontDescriptor",
        "reusable CIDFont descriptor",
      );
      if (!descriptorRef) continue;
      const descriptorObject = parseIndirectObject(
        input.sourceBytes,
        parsed,
        descriptorRef.objectNumber,
        descriptorRef.generation,
      );
      const descriptorDictionary = topLevelDictionarySource(
        asLatin1(descriptorObject.bytes),
        "reusable FontDescriptor",
      );
      const fontFileRef = oneTopLevelReference(
        descriptorDictionary,
        "FontFile2",
        "reusable TrueType font file",
      );
      if (!fontFileRef) continue;
      const fontFileObject = parseIndirectObject(
        input.sourceBytes,
        parsed,
        fontFileRef.objectNumber,
        fontFileRef.generation,
      );
      if (!fontFileObject.streamBytes) continue;
      // Rebuild against the program this document already installed. A
      // narrowed program only matches byte for byte when it keeps the glyphs
      // the page can already draw, and matching is what lets this edit reuse
      // the installed font instead of installing a second one.
      const fontPackage = await buildPdfNativeFontExtensionPackage({
        fontBytes: input.fontBytes,
        text: input.replacementText,
        resourceName: entry.resourceName,
        firstObjectNumber: entry.target.objectNumber,
        retainedProgram: fontFileObject.streamBytes,
      });
      const expectedToUnicode = fontPackage.objects.find(
        (object) =>
          object.objectNumber === toUnicodeRef.objectNumber &&
          object.generation === toUnicodeRef.generation,
      );
      if (!expectedToUnicode) continue;
      const changedObject = fontPackage.objects.find((expected) => {
        if (expected.objectNumber === toUnicodeRef.objectNumber) return false;
        const actual = parseIndirectObject(
            input.sourceBytes,
            parsed,
            expected.objectNumber,
            expected.generation,
          );
          return !bytesEqual(actual.bytes, expected.bytes);
        });
      if (
        fontPackage.font.objectNumber !== entry.target.objectNumber ||
        changedObject
      ) {
        continue;
      }
      const currentToUnicode = parseIndirectObject(
        input.sourceBytes,
        parsed,
        toUnicodeRef.objectNumber,
        toUnicodeRef.generation,
      );
      const replacementToUnicode = parseDirectObjectAtOffset(
        expectedToUnicode.bytes,
        0,
        {
          objectNumber: expectedToUnicode.objectNumber,
          generation: expectedToUnicode.generation,
        },
      );
      if (
        currentToUnicode.streamBytes === null ||
        replacementToUnicode.streamBytes === null
      ) {
        continue;
      }
      return Object.freeze({
        sourceSha256,
        pageNumber,
        resourceName: entry.resourceName,
        font: fontPackage.font,
        fontObject: Object.freeze({
          objectNumber: fontObject.objectNumber,
          generation: fontObject.generation,
          objectSha256: await sha256PdfBytes(fontObject.bytes),
          objectBytes: Uint8Array.from(fontObject.bytes),
        }),
        resources: Object.freeze({
          page: Object.freeze({ ...page.page }),
          font: Object.freeze({ ...fontPackage.font }),
          descendantFont: Object.freeze({ ...descendantRef }),
          descriptor: Object.freeze({ ...descriptorRef }),
          fontFile: Object.freeze({ ...fontFileRef }),
          toUnicode: Object.freeze({ ...toUnicodeRef }),
        }),
        fontPackage,
        sharedObjects: Object.freeze(
          fontPackage.objects
            .filter(
              (object) => object.objectNumber !== toUnicodeRef.objectNumber,
            )
            .map((object) =>
              Object.freeze({
                objectNumber: object.objectNumber,
                generation: object.generation,
                sha256: object.sha256,
              }),
            ),
        ),
        currentToUnicode: Object.freeze({
          objectNumber: currentToUnicode.objectNumber,
          generation: currentToUnicode.generation,
          objectSha256: await sha256PdfBytes(currentToUnicode.bytes),
          streamSha256: await sha256PdfBytes(currentToUnicode.streamBytes),
          streamBytes: Uint8Array.from(currentToUnicode.streamBytes),
          objectBytes: Uint8Array.from(currentToUnicode.bytes),
        }),
        replacementToUnicode: Object.freeze({
          objectNumber: replacementToUnicode.objectNumber,
          generation: replacementToUnicode.generation as 0,
          objectSha256: await sha256PdfBytes(replacementToUnicode.bytes),
          streamSha256: await sha256PdfBytes(
            replacementToUnicode.streamBytes,
          ),
          streamBytes: Uint8Array.from(replacementToUnicode.streamBytes),
          objectBytes: Uint8Array.from(replacementToUnicode.bytes),
        }),
      });
    } catch (error) {
      if (!(error instanceof PdfNativeTransactionError)) throw error;
    }
  }
  return null;
}

function rewriteFontAliasObjectIdentity(
  bytes: Uint8Array,
  current: PdfNativeIndirectObjectRef,
  replacement: PdfNativeIndirectObjectRef,
  label: string,
): Uint8Array {
  const source = asLatin1(bytes);
  const header = new RegExp(
    `^${current.objectNumber}\\s+${current.generation}\\s+obj\\b`,
  );
  if (!header.test(source)) {
    fail("precondition-mismatch", `${label} identity changed.`);
  }
  return latin1Bytes(
    source.replace(
      header,
      `${replacement.objectNumber} ${replacement.generation} obj`,
    ),
  );
}

/** Build a small Type0+CMap alias while retaining the proven CIDFont program. */
export async function buildMraiNativeFontAliasPackage(input: {
  readonly reusable: PdfNativeReusablePageFontExtension;
  readonly resourceName: string;
  readonly firstObjectNumber: number;
  readonly replacementText: string;
}): Promise<PdfNativeFontAliasPackage> {
  if (
    !PDF_RESOURCE_NAME.test(input.resourceName) ||
    !Number.isSafeInteger(input.firstObjectNumber) ||
    input.firstObjectNumber < 1 ||
    input.firstObjectNumber > 2_147_483_600
  ) {
    fail("invalid-request", "Native font alias identity is invalid.");
  }
  const font = Object.freeze({
    objectNumber: input.firstObjectNumber,
    generation: 0 as const,
  });
  const toUnicode = Object.freeze({
    objectNumber: input.firstObjectNumber + 1,
    generation: 0 as const,
  });
  let fontBytes = rewriteFontAliasObjectIdentity(
    input.reusable.fontObject.objectBytes,
    input.reusable.fontObject,
    font,
    "Reusable Type0 font",
  );
  const fontSource = asLatin1(fontBytes);
  const mappingPattern = new RegExp(
    `/ToUnicode\\s+${input.reusable.currentToUnicode.objectNumber}\\s+${input.reusable.currentToUnicode.generation}\\s+R\\b`,
    "g",
  );
  if ([...fontSource.matchAll(mappingPattern)].length !== 1) {
    fail("precondition-mismatch", "Reusable ToUnicode mapping changed.");
  }
  fontBytes = latin1Bytes(
    fontSource.replace(
      mappingPattern,
      `/ToUnicode ${toUnicode.objectNumber} ${toUnicode.generation} R`,
    ),
  );
  const characters = [...input.replacementText];
  const encoded = input.reusable.fontPackage.encodedHex.match(/.{4}/gu) ?? [];
  if (
    characters.length !== encoded.length ||
    characters.length < 1 ||
    characters.length > 2_000
  ) {
    fail("invalid-request", "Native font alias glyph mapping is invalid.");
  }
  // One entry per code: a text that repeats a letter encodes its glyph again,
  // and a ToUnicode naming a code twice is ambiguous to every strict reader
  // (replay refused every alias of a text with a repeated letter). A code two
  // different characters share cannot be told apart at all.
  const unicodeByCode = new Map<string, string>();
  characters.forEach((character, index) => {
    const unicode = [...character]
      .flatMap((scalar) => {
        const point = scalar.codePointAt(0)!;
        if (point <= 0xffff) return [point];
        const adjusted = point - 0x10000;
        return [0xd800 + (adjusted >> 10), 0xdc00 + (adjusted & 0x3ff)];
      })
      .map((value) => value.toString(16).padStart(4, "0").toUpperCase())
      .join("");
    const code = encoded[index].toUpperCase();
    const known = unicodeByCode.get(code);
    if (known !== undefined && known !== unicode) {
      fail("invalid-request", "Native font alias glyph mapping is ambiguous.");
    }
    unicodeByCode.set(code, unicode);
  });
  const mappings = [...unicodeByCode].map(([code, unicode]) => `<${code}> <${unicode}>`);
  const cmap = latin1Bytes(
    [
      "/CIDInit /ProcSet findresource begin",
      "12 dict begin",
      "begincmap",
      "/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def",
      "/CMapName /Adobe-Identity-UCS def",
      "/CMapType 2 def",
      "1 begincodespacerange",
      "<0000> <FFFF>",
      "endcodespacerange",
      `${mappings.length} beginbfchar`,
      ...mappings,
      "endbfchar",
      "endcmap",
      "CMapName currentdict /CMap defineresource pop",
      "end",
      "end",
      "",
    ].join("\n"),
  );
  const toUnicodeBytes = concatBytes([
    ascii(`${toUnicode.objectNumber} 0 obj\n<< /Length ${cmap.length} >>\nstream\n`),
    cmap,
    ascii("endstream\nendobj"),
  ]);
  const objects = Object.freeze([
    Object.freeze({
      ...font,
      sha256: await sha256PdfBytes(fontBytes),
      bytes: fontBytes,
    }),
    Object.freeze({
      ...toUnicode,
      sha256: await sha256PdfBytes(toUnicodeBytes),
      bytes: toUnicodeBytes,
    }),
  ]);
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_FONT_ALIAS_PACKAGE_SCHEMA_VERSION,
    fontSha256: PDF_UNICODE_FALLBACK_FONT_SHA256,
    resourceName: input.resourceName,
    font,
    sourceFont: Object.freeze({
      objectNumber: input.reusable.fontObject.objectNumber,
      generation: input.reusable.fontObject.generation,
      expectedObjectSha256: input.reusable.fontObject.objectSha256,
    }),
    sourceToUnicode: Object.freeze({
      objectNumber: input.reusable.currentToUnicode.objectNumber,
      generation: input.reusable.currentToUnicode.generation,
      expectedObjectSha256: input.reusable.currentToUnicode.objectSha256,
    }),
    encodedHex: input.reusable.fontPackage.encodedHex,
    widthAt1000: input.reusable.fontPackage.widthAt1000,
    objectByteLength: objects.reduce(
      (total, object) => total + object.bytes.length,
      0,
    ),
    objects,
  });
}

function fontExtensionHistoryRevisionBytes(
  object: ParsedIndirectObject,
): Uint8Array {
  return object.storage === "direct"
    ? Uint8Array.from(object.bytes)
    : concatBytes([
        ascii(`${object.objectNumber} ${object.generation} obj\n`),
        object.bytes,
        ascii("\nendobj"),
      ]);
}

/**
 * Append one exact two-object undo/redo revision for a previously verified
 * page-local font extension. The font/resource objects stay byte-identical;
 * switching the page and content-stream objects makes them reachable or
 * unreachable without rewriting the original package.
 */
export async function planMraiNativePageFontExtensionHistory(input: {
  readonly sourceBytes: Uint8Array;
  readonly originSourceBytes: Uint8Array;
  readonly originPrepared: PdfNativePreparedContentStreamTransaction;
  readonly direction: "undo" | "redo";
}): Promise<PdfNativePreparedFontExtensionHistory> {
  const { originPrepared } = input;
  const mutation = originPrepared.request.pageFontExtensionMutation;
  const receipt = originPrepared.pageFontExtensionMutation;
  if (
    !mutation ||
    !receipt ||
    originPrepared.sourceLength !== input.originSourceBytes.length ||
    (await sha256PdfBytes(input.originSourceBytes)) !==
      originPrepared.request.identity.sourceSha256 ||
    (await sha256PdfBytes(originPrepared.candidateBytes)) !==
      originPrepared.candidateSha256 ||
    input.sourceBytes.length < originPrepared.candidateBytes.length ||
    !bytesEqual(
      input.sourceBytes.slice(0, originPrepared.candidateBytes.length),
      originPrepared.candidateBytes,
    )
  ) {
    fail("precondition-mismatch", "Font extension history origin changed.");
  }

  const original = parsePdf(input.originSourceBytes, false);
  const extended = parsePdf(originPrepared.candidateBytes, true);
  const current = parsePdf(input.sourceBytes, false);
  if (
    original.current.offset !== originPrepared.sourceXrefOffset ||
    extended.current.offset !== originPrepared.incrementalXrefOffset ||
    current.sections.length >= 32
  ) {
    fail("precondition-mismatch", "Font extension history chain changed.");
  }

  const targetRef = originPrepared.request.identity;
  const pageRef = mutation.page;
  const originalTarget = parseIndirectObject(
    input.originSourceBytes,
    original,
    targetRef.objectNumber,
    targetRef.generation,
  );
  const extendedTarget = parseIndirectObject(
    originPrepared.candidateBytes,
    extended,
    targetRef.objectNumber,
    targetRef.generation,
  );
  const currentTarget = parseIndirectObject(
    input.sourceBytes,
    current,
    targetRef.objectNumber,
    targetRef.generation,
  );
  const originalPage = parseIndirectObject(
    input.originSourceBytes,
    original,
    pageRef.objectNumber,
    pageRef.generation,
  );
  const extendedPage = parseIndirectObject(
    originPrepared.candidateBytes,
    extended,
    pageRef.objectNumber,
    pageRef.generation,
  );
  const currentPage = parseIndirectObject(
    input.sourceBytes,
    current,
    pageRef.objectNumber,
    pageRef.generation,
  );
  if (
    originalTarget.encodedStreamBytes === null ||
    extendedTarget.encodedStreamBytes === null ||
    currentTarget.encodedStreamBytes === null
  ) {
    fail("unsupported-document", "Font extension history stream is invalid.");
  }

  const originalTargetBytes = fontExtensionHistoryRevisionBytes(originalTarget);
  const extendedTargetBytes = fontExtensionHistoryRevisionBytes(extendedTarget);
  const originalPageBytes = fontExtensionHistoryRevisionBytes(originalPage);
  const extendedPageBytes = fontExtensionHistoryRevisionBytes(extendedPage);
  const currentTargetBytes = fontExtensionHistoryRevisionBytes(currentTarget);
  const currentPageBytes = fontExtensionHistoryRevisionBytes(currentPage);
  const expectedTargetBytes =
    input.direction === "undo" ? extendedTargetBytes : originalTargetBytes;
  const expectedPageBytes =
    input.direction === "undo" ? extendedPageBytes : originalPageBytes;
  const desiredTargetBytes =
    input.direction === "undo" ? originalTargetBytes : extendedTargetBytes;
  const desiredPageBytes =
    input.direction === "undo" ? originalPageBytes : extendedPageBytes;
  if (
    !bytesEqual(currentTargetBytes, expectedTargetBytes) ||
    !bytesEqual(currentPageBytes, expectedPageBytes)
  ) {
    fail("precondition-mismatch", "Font extension history side changed.");
  }

  for (const item of [...receipt.fontObjects, receipt.resources]) {
    const originObject = parseIndirectObject(
      originPrepared.candidateBytes,
      extended,
      item.objectNumber,
      item.generation,
    );
    const currentObject = parseIndirectObject(
      input.sourceBytes,
      current,
      item.objectNumber,
      item.generation,
    );
    if (
      (await sha256PdfBytes(originObject.bytes)) !== item.sha256 ||
      !bytesEqual(currentObject.bytes, originObject.bytes)
    ) {
      fail("precondition-mismatch", "Font extension resource changed.");
    }
  }

  const revisions: IncrementalObjectRevision[] = [
    {
      objectNumber: targetRef.objectNumber,
      generation: targetRef.generation,
      bytes: desiredTargetBytes,
    },
    {
      objectNumber: pageRef.objectNumber,
      generation: pageRef.generation,
      bytes: desiredPageBytes,
    },
  ].sort(
    (left, right) =>
      left.objectNumber - right.objectNumber ||
      left.generation - right.generation,
  );
  const incremental = makeMultiObjectIncrementalRevision(
    input.sourceBytes,
    current.current,
    revisions,
    { min: 2, max: 2, label: "Font extension history" },
  );
  const candidate = parsePdf(incremental.candidateBytes, true);
  const revisedObjects = await Promise.all(
    revisions.map(async (revision) => {
      const offset = incremental.offsets.get(
        `${revision.objectNumber}:${revision.generation}`,
      );
      if (offset === undefined) {
        fail("verification-failed", "Font history receipt is incomplete.");
      }
      const reopened = parseIndirectObject(
        incremental.candidateBytes,
        candidate,
        revision.objectNumber,
        revision.generation,
      );
      if (!bytesEqual(reopened.bytes, revision.bytes)) {
        fail("verification-failed", "Font history object failed reopen.");
      }
      return Object.freeze({
        objectNumber: revision.objectNumber,
        generation: revision.generation,
        offset,
        sha256: await sha256PdfBytes(revision.bytes),
      });
    }),
  );
  const desiredTarget = parseIndirectObject(
    incremental.candidateBytes,
    candidate,
    targetRef.objectNumber,
    targetRef.generation,
  );
  if (
    desiredTarget.encodedStreamBytes === null ||
    !bytesEqual(
      incremental.candidateBytes.slice(0, input.sourceBytes.length),
      input.sourceBytes,
    )
  ) {
    fail("verification-failed", "Font extension history custody failed.");
  }
  const state = async (
    targetBytes: Uint8Array,
    streamBytes: Uint8Array,
    pageBytes: Uint8Array,
  ): Promise<PdfNativeFontExtensionHistoryObjectState> =>
    Object.freeze({
      targetObjectSha256: await sha256PdfBytes(targetBytes),
      targetStreamSha256: await sha256PdfBytes(streamBytes),
      pageObjectSha256: await sha256PdfBytes(pageBytes),
    });
  const originalState = await state(
    originalTargetBytes,
    originalTarget.encodedStreamBytes,
    originalPageBytes,
  );
  const extendedState = await state(
    extendedTargetBytes,
    extendedTarget.encodedStreamBytes,
    extendedPageBytes,
  );
  return Object.freeze({
    direction: input.direction,
    target: Object.freeze({ ...targetRef }),
    page: Object.freeze({
      objectNumber: pageRef.objectNumber,
      generation: pageRef.generation,
    }),
    current: input.direction === "undo" ? extendedState : originalState,
    desired: input.direction === "undo" ? originalState : extendedState,
    sourceLength: input.sourceBytes.length,
    sourceXrefOffset: current.current.offset,
    incrementalXrefOffset: incremental.incrementalXrefOffset,
    changedObjectOffset: incremental.offsets.get(
      `${targetRef.objectNumber}:${targetRef.generation}`,
    )!,
    changedObjectSha256: await sha256PdfBytes(desiredTargetBytes),
    changedStreamSha256: await sha256PdfBytes(
      desiredTarget.encodedStreamBytes,
    ),
    candidateSha256: await sha256PdfBytes(incremental.candidateBytes),
    candidateBytes: incremental.candidateBytes,
    revisedObjects: Object.freeze(revisedObjects),
  });
}

/** Append one exact target-only undo/redo revision for a reused page font. */
export async function planMraiNativeFontReuseHistory(input: {
  readonly sourceBytes: Uint8Array;
  readonly originSourceBytes: Uint8Array;
  readonly originPrepared: PdfNativePreparedContentStreamTransaction;
  readonly page: PdfNativeIndirectObjectRef;
  readonly direction: "undo" | "redo";
}): Promise<PdfNativePreparedFontExtensionHistory> {
  const { originPrepared } = input;
  if (
    originPrepared.request.pageFontExtensionMutation ||
    originPrepared.pageFontExtensionMutation ||
    originPrepared.sourceLength !== input.originSourceBytes.length ||
    (await sha256PdfBytes(input.originSourceBytes)) !==
      originPrepared.request.identity.sourceSha256 ||
    (await sha256PdfBytes(originPrepared.candidateBytes)) !==
      originPrepared.candidateSha256 ||
    input.sourceBytes.length < originPrepared.candidateBytes.length ||
    !bytesEqual(
      input.sourceBytes.slice(0, originPrepared.candidateBytes.length),
      originPrepared.candidateBytes,
    )
  ) {
    fail("precondition-mismatch", "Font reuse history origin changed.");
  }
  const original = parsePdf(input.originSourceBytes, false);
  const edited = parsePdf(originPrepared.candidateBytes, true);
  const current = parsePdf(input.sourceBytes, false);
  if (
    original.current.offset !== originPrepared.sourceXrefOffset ||
    edited.current.offset !== originPrepared.incrementalXrefOffset ||
    current.sections.length >= 32
  ) {
    fail("precondition-mismatch", "Font reuse history chain changed.");
  }
  const target = originPrepared.request.identity;
  const originalTarget = parseIndirectObject(
    input.originSourceBytes,
    original,
    target.objectNumber,
    target.generation,
  );
  const editedTarget = parseIndirectObject(
    originPrepared.candidateBytes,
    edited,
    target.objectNumber,
    target.generation,
  );
  const currentTarget = parseIndirectObject(
    input.sourceBytes,
    current,
    target.objectNumber,
    target.generation,
  );
  const originalPage = parseIndirectObject(
    input.originSourceBytes,
    original,
    input.page.objectNumber,
    input.page.generation,
  );
  const editedPage = parseIndirectObject(
    originPrepared.candidateBytes,
    edited,
    input.page.objectNumber,
    input.page.generation,
  );
  const currentPage = parseIndirectObject(
    input.sourceBytes,
    current,
    input.page.objectNumber,
    input.page.generation,
  );
  if (
    originalTarget.encodedStreamBytes === null ||
    editedTarget.encodedStreamBytes === null ||
    currentTarget.encodedStreamBytes === null ||
    !bytesEqual(originalPage.bytes, editedPage.bytes) ||
    !bytesEqual(originalPage.bytes, currentPage.bytes)
  ) {
    fail("precondition-mismatch", "Font reuse history ownership changed.");
  }
  const originalTargetBytes = fontExtensionHistoryRevisionBytes(originalTarget);
  const editedTargetBytes = fontExtensionHistoryRevisionBytes(editedTarget);
  const currentTargetBytes = fontExtensionHistoryRevisionBytes(currentTarget);
  const expectedTargetBytes =
    input.direction === "undo" ? editedTargetBytes : originalTargetBytes;
  const desiredTargetBytes =
    input.direction === "undo" ? originalTargetBytes : editedTargetBytes;
  if (!bytesEqual(currentTargetBytes, expectedTargetBytes)) {
    fail("precondition-mismatch", "Font reuse history side changed.");
  }
  for (const witness of originPrepared.request.untouchedObjects ?? []) {
    const originObject = parseIndirectObject(
      originPrepared.candidateBytes,
      edited,
      witness.objectNumber,
      witness.generation,
    );
    const currentObject = parseIndirectObject(
      input.sourceBytes,
      current,
      witness.objectNumber,
      witness.generation,
    );
    if (
      (await sha256PdfBytes(originObject.bytes)) !==
        witness.expectedObjectSha256 ||
      !bytesEqual(currentObject.bytes, originObject.bytes)
    ) {
      fail("precondition-mismatch", "Font reuse resource changed.");
    }
  }
  const incremental = makeMultiObjectIncrementalRevision(
    input.sourceBytes,
    current.current,
    [
      {
        objectNumber: target.objectNumber,
        generation: target.generation,
        bytes: desiredTargetBytes,
      },
    ],
    { min: 1, max: 1, label: "Font reuse history" },
  );
  const candidate = parsePdf(incremental.candidateBytes, true);
  const offset = incremental.offsets.get(
    `${target.objectNumber}:${target.generation}`,
  );
  if (offset === undefined) {
    fail("verification-failed", "Font reuse history receipt is incomplete.");
  }
  const reopened = parseIndirectObject(
    incremental.candidateBytes,
    candidate,
    target.objectNumber,
    target.generation,
  );
  if (
    reopened.encodedStreamBytes === null ||
    !bytesEqual(reopened.bytes, desiredTargetBytes) ||
    !bytesEqual(
      incremental.candidateBytes.slice(0, input.sourceBytes.length),
      input.sourceBytes,
    )
  ) {
    fail("verification-failed", "Font reuse history failed reopen.");
  }
  const pageSha256 = await sha256PdfBytes(
    fontExtensionHistoryRevisionBytes(originalPage),
  );
  const state = async (
    objectBytes: Uint8Array,
    streamBytes: Uint8Array,
  ): Promise<PdfNativeFontExtensionHistoryObjectState> =>
    Object.freeze({
      targetObjectSha256: await sha256PdfBytes(objectBytes),
      targetStreamSha256: await sha256PdfBytes(streamBytes),
      pageObjectSha256: pageSha256,
    });
  const originalState = await state(
    originalTargetBytes,
    originalTarget.encodedStreamBytes,
  );
  const editedState = await state(
    editedTargetBytes,
    editedTarget.encodedStreamBytes,
  );
  return Object.freeze({
    direction: input.direction,
    target: Object.freeze({ ...target }),
    page: Object.freeze({ ...input.page }),
    current: input.direction === "undo" ? editedState : originalState,
    desired: input.direction === "undo" ? originalState : editedState,
    sourceLength: input.sourceBytes.length,
    sourceXrefOffset: current.current.offset,
    incrementalXrefOffset: incremental.incrementalXrefOffset,
    changedObjectOffset: offset,
    changedObjectSha256: await sha256PdfBytes(desiredTargetBytes),
    changedStreamSha256: await sha256PdfBytes(reopened.encodedStreamBytes),
    candidateSha256: await sha256PdfBytes(incremental.candidateBytes),
    candidateBytes: incremental.candidateBytes,
    revisedObjects: Object.freeze([
      Object.freeze({
        objectNumber: target.objectNumber,
        generation: target.generation,
        offset,
        sha256: await sha256PdfBytes(desiredTargetBytes),
      }),
    ]),
  });
}

/**
 * The undo/redo revision of a font-extension or font-reuse edit whose own
 * revisions a source consolidation compacted away: the same object swap the
 * two planners above write, its original and extended objects read from the
 * evidence the consolidation record kept for the edit (each object as a
 * history revision writes it, the digests of the stream each target carries,
 * and the objects the edit installed that must still read as it wrote them)
 * instead of from the edit's revisions. Mirror of the server's
 * _replay_native_font_extension_history_from_evidence.
 */
export async function planMraiNativeFontHistoryFromEvidence(input: {
  readonly sourceBytes: Uint8Array;
  readonly kept: Readonly<{
    originKind: "extension" | "reuse";
    originSourceSha256: string;
    target: PdfNativeIndirectObjectRef;
    page: PdfNativeIndirectObjectRef;
    originalTarget: Uint8Array;
    extendedTarget: Uint8Array;
    originalPage: Uint8Array;
    extendedPage: Uint8Array;
    originalTargetStreamSha256: string;
    extendedTargetStreamSha256: string;
    protected: readonly Readonly<{ objectNumber: number; generation: number; sha256: string }>[];
  }>;
  readonly direction: "undo" | "redo";
}): Promise<PdfNativePreparedFontExtensionHistory> {
  const { kept } = input;
  const undo = input.direction === "undo";
  const reuse = kept.originKind === "reuse";
  const current = parsePdf(input.sourceBytes, false);
  if (current.sections.length >= 32) {
    fail("precondition-mismatch", "Font history chain changed.");
  }
  const currentTarget = parseIndirectObject(
    input.sourceBytes, current, kept.target.objectNumber, kept.target.generation,
  );
  const currentPage = parseIndirectObject(
    input.sourceBytes, current, kept.page.objectNumber, kept.page.generation,
  );
  if (
    currentTarget.encodedStreamBytes === null ||
    !bytesEqual(fontExtensionHistoryRevisionBytes(currentTarget), undo ? kept.extendedTarget : kept.originalTarget) ||
    !bytesEqual(fontExtensionHistoryRevisionBytes(currentPage), undo ? kept.extendedPage : kept.originalPage)
  ) {
    fail("precondition-mismatch", "Font history side changed.");
  }
  for (const item of kept.protected) {
    const installed = parseIndirectObject(input.sourceBytes, current, item.objectNumber, item.generation);
    if ((await sha256PdfBytes(installed.bytes)) !== item.sha256) {
      fail("precondition-mismatch", "Font extension resource changed.");
    }
  }
  const desiredTargetBytes = undo ? kept.originalTarget : kept.extendedTarget;
  const desiredStreamSha256 = undo ? kept.originalTargetStreamSha256 : kept.extendedTargetStreamSha256;
  const revisions: IncrementalObjectRevision[] = [
    { objectNumber: kept.target.objectNumber, generation: kept.target.generation, bytes: desiredTargetBytes },
    ...(reuse
      ? []
      : [{
          objectNumber: kept.page.objectNumber,
          generation: kept.page.generation,
          bytes: undo ? kept.originalPage : kept.extendedPage,
        }]),
  ].sort((left, right) => left.objectNumber - right.objectNumber || left.generation - right.generation);
  const incremental = makeMultiObjectIncrementalRevision(
    input.sourceBytes,
    current.current,
    revisions,
    reuse
      ? { min: 1, max: 1, label: "Font reuse history" }
      : { min: 2, max: 2, label: "Font extension history" },
  );
  const candidate = parsePdf(incremental.candidateBytes, true);
  const revisedObjects = await Promise.all(
    revisions.map(async (revision) => {
      const offset = incremental.offsets.get(`${revision.objectNumber}:${revision.generation}`);
      if (offset === undefined) fail("verification-failed", "Font history receipt is incomplete.");
      const reopened = parseIndirectObject(
        incremental.candidateBytes, candidate, revision.objectNumber, revision.generation,
      );
      if (!bytesEqual(reopened.bytes, revision.bytes)) {
        fail("verification-failed", "Font history object failed reopen.");
      }
      return Object.freeze({
        objectNumber: revision.objectNumber,
        generation: revision.generation,
        offset,
        sha256: await sha256PdfBytes(revision.bytes),
      });
    }),
  );
  const desiredTarget = parseIndirectObject(
    incremental.candidateBytes, candidate, kept.target.objectNumber, kept.target.generation,
  );
  if (
    desiredTarget.encodedStreamBytes === null ||
    (await sha256PdfBytes(desiredTarget.encodedStreamBytes)) !== desiredStreamSha256 ||
    !bytesEqual(incremental.candidateBytes.slice(0, input.sourceBytes.length), input.sourceBytes)
  ) {
    fail("verification-failed", "Font history custody failed.");
  }
  const originalState: PdfNativeFontExtensionHistoryObjectState = Object.freeze({
    targetObjectSha256: await sha256PdfBytes(kept.originalTarget),
    targetStreamSha256: kept.originalTargetStreamSha256,
    pageObjectSha256: await sha256PdfBytes(kept.originalPage),
  });
  const extendedState: PdfNativeFontExtensionHistoryObjectState = Object.freeze({
    targetObjectSha256: await sha256PdfBytes(kept.extendedTarget),
    targetStreamSha256: kept.extendedTargetStreamSha256,
    pageObjectSha256: await sha256PdfBytes(kept.extendedPage),
  });
  return Object.freeze({
    direction: input.direction,
    target: Object.freeze({
      sourceSha256: kept.originSourceSha256,
      objectNumber: kept.target.objectNumber,
      generation: kept.target.generation,
    }),
    page: Object.freeze({ objectNumber: kept.page.objectNumber, generation: kept.page.generation }),
    current: undo ? extendedState : originalState,
    desired: undo ? originalState : extendedState,
    sourceLength: input.sourceBytes.length,
    sourceXrefOffset: current.current.offset,
    incrementalXrefOffset: incremental.incrementalXrefOffset,
    changedObjectOffset: incremental.offsets.get(`${kept.target.objectNumber}:${kept.target.generation}`)!,
    changedObjectSha256: await sha256PdfBytes(desiredTargetBytes),
    changedStreamSha256: desiredStreamSha256,
    candidateSha256: await sha256PdfBytes(incremental.candidateBytes),
    candidateBytes: incremental.candidateBytes,
    revisedObjects: Object.freeze(revisedObjects),
  });
}

export function nextMraiNativeIndirectObjectNumber(
  sourceBytes: Uint8Array,
): number {
  const parsed = parsePdf(sourceBytes, false);
  const latest = latestEntries(parsed);
  const next = Math.max(parsed.current.size, Math.max(0, ...latest.keys()) + 1);
  if (!Number.isSafeInteger(next) || next < 1 || next > 2_147_483_600) {
    fail("unsupported-document", "PDF object allocation is unavailable.");
  }
  return next;
}

/** The dictionary of one directly stored indirect object as written, read
 * without decoding its stream: what a footprint or subtype question needs
 * of a picture whose filter this engine cannot decode. Null when the object
 * is not stored directly or its dictionary cannot be delimited. */
/** The parsed structure of one source, kept while the same bytes are asked
 * about again: a plan over a long book inspects an object per page, and
 * the xref of a 20 MB file is not read over for each. */
const parsedSourceStructures = new WeakMap<Uint8Array, ParsedPdf>();

function parseSourceStructure(sourceBytes: Uint8Array): ParsedPdf {
  const cached = parsedSourceStructures.get(sourceBytes);
  if (cached) return cached;
  const parsed = parsePdf(sourceBytes, false);
  parsedSourceStructures.set(sourceBytes, parsed);
  return parsed;
}

export function readMraiNativeObjectDictionarySource(
  sourceBytes: Uint8Array,
  target: Readonly<{ objectNumber: number; generation: number }>,
): string | null {
  try {
    assertObjectReference(target.objectNumber, target.generation);
    const parsed = parseSourceStructure(sourceBytes);
    const entry = resolveXrefEntry(parsed, target.objectNumber, target.generation);
    if (entry.kind !== "direct") return null;
    const window = asLatin1(
      sourceBytes.subarray(entry.offset, Math.min(sourceBytes.length, entry.offset + 65_536)),
    );
    const header = /^[\s\x00]*\d+\s+\d+\s+obj\b/.exec(window);
    if (!header) return null;
    const start = window.indexOf("<<", header[0].length);
    if (start < 0 || /\S/.test(window.slice(header[0].length, start))) return null;
    return window.slice(start, dictionaryEnd(window, start));
  } catch (error) {
    if (!(error instanceof PdfNativeTransactionError)) throw error;
    return null;
  }
}

/** One source's digest, computed once per buffer: a plan inspects an
 * object per page of a 1200-page book, and each inspection names the
 * source it read from; hashing 23 MB again every time was most of its
 * planning time. Sources are never written in place. */
const sourceDigests = new WeakMap<Uint8Array, Promise<string>>();
function sourceDigest(sourceBytes: Uint8Array): Promise<string> {
  let digest = sourceDigests.get(sourceBytes);
  if (!digest) {
    digest = sha256PdfBytes(sourceBytes);
    sourceDigests.set(sourceBytes, digest);
  }
  return digest;
}

export async function inspectMraiNativeObject(
  sourceBytes: Uint8Array,
  target: Readonly<{ objectNumber: number; generation: number }>,
): Promise<PdfNativeObjectInspection> {
  assertObjectReference(target.objectNumber, target.generation);
  const parsed = parseSourceStructure(sourceBytes);
  const sourceSha256 = await sourceDigest(sourceBytes);
  return inspectNativeObjectInReadSource(sourceBytes, parsed, sourceSha256, target);
}

async function inspectNativeObjectInReadSource(
  sourceBytes: Uint8Array, parsed: ParsedPdf, sourceSha256: string,
  target: Readonly<{ objectNumber: number; generation: number }>,
  readContext?: NativeObjectReadContext,
): Promise<PdfNativeObjectInspection> {
  assertObjectReference(target.objectNumber, target.generation);
  const object = parseIndirectObject(
    sourceBytes,
    parsed,
    target.objectNumber,
    target.generation,
    readContext,
  );
  return Object.freeze({
    identity: Object.freeze({
      sourceSha256,
      objectNumber: target.objectNumber,
      generation: target.generation,
    }),
    objectSha256: await sha256PdfBytes(object.bytes),
    objectBytes: Uint8Array.from(object.bytes),
    streamEncoding: object.streamEncoding,
    streamSha256: object.encodedStreamBytes
      ? await sha256PdfBytes(object.encodedStreamBytes)
      : null,
    streamBytes: object.streamBytes
      ? Uint8Array.from(object.streamBytes)
      : null,
  });
}

/**
 * A tagged document's structure tree names the stream each of its leaves was
 * written in: a marked-content reference dictionary (`/Type /MCR`) points at
 * that stream with `/Stm`, and at the object that owns it with `/StmOwn`
 * (ISO 32000-1 Table 324). Neither draws anything, so neither is a second
 * place the stream is used -- but a count that reads them as one made every
 * form a tagged page draws once look shared, and nothing drawn inside such a
 * form could be addressed at all. Those two values are blanked, in place and
 * byte for byte, before the references are counted; every other reference to
 * the same object, including one from the very same structure element, still
 * counts.
 */
const MARKED_CONTENT_REFERENCE_DICTIONARY = /<<(?:(?!<<|>>)[\s\S])*>>/g;
const MARKED_CONTENT_STREAM_OWNER = /(\/(?:Stm|StmOwn)\s*)(\d+\s+\d+\s+R)\b/g;

function withoutMarkedContentStreamOwners(dictionary: string): string {
  if (!dictionary.includes("/MCR")) return dictionary;
  return dictionary.replace(MARKED_CONTENT_REFERENCE_DICTIONARY, (span) =>
    /\/Type\s*\/MCR\b/.test(span)
      ? span.replace(
          MARKED_CONTENT_STREAM_OWNER,
          (_match, key: string, target: string) => `${key}${" ".repeat(target.length)}`,
        )
      : span,
  );
}

/**
 * Count dictionary-level references to one pinned indirect object across the
 * current object graph. Stream payload bytes are deliberately excluded so a
 * mutable Form XObject can be refused unless it has one unambiguous owner.
 */
export function countMraiNativeIndirectObjectReferences(
  sourceBytes: Uint8Array,
  target: Readonly<{ objectNumber: number; generation: number }>,
): number {
  return countMraiNativeIndirectObjectReferencesBatch(sourceBytes, [target])[0];
}

/**
 * Count many object references from one authenticated object-graph parse.
 * The result is aligned with `targets`, including duplicate target rows.
 */
export function countMraiNativeIndirectObjectReferencesBatch(
  sourceBytes: Uint8Array,
  targets: readonly Readonly<{
    objectNumber: number;
    generation: number;
  }>[],
): readonly number[] {
  if (!Array.isArray(targets) || targets.length > MAX_REFERENCE_COUNT_TARGETS) {
    fail("invalid-request", "PDF reference-count target set is invalid.");
  }
  const keyFor = (target: Readonly<{ objectNumber: number; generation: number }>) =>
    `${target.objectNumber}:${target.generation}`;
  const counts = new Map<string, number>();
  for (const target of targets) {
    assertObjectReference(target.objectNumber, target.generation);
    counts.set(keyFor(target), 0);
  }
  if (targets.length === 0) return Object.freeze([]);
  const parsed = parsePdf(sourceBytes, false);
  const reference = /(?:^|[^0-9])([1-9][0-9]{0,9})\s+([0-9]{1,5})\s+R\b/g;
  // Decode the source once for this synchronous traversal. Do not retain
  // caller-owned bytes or share this cache with a later read or writer.
  const readContext = {
    source: asLatin1(sourceBytes),
    members: new Map<number, readonly ObjectStreamMember[]>(),
  };
  for (const entry of latestEntries(parsed).values()) {
    if (!entry.inUse) continue;
    // Only the dictionaries hold references; a picture this engine cannot
    // decode (a TIFF predictor, say) still says what it points at, so its
    // dictionary is read on its own rather than making the whole count
    // refuse. An object whose dictionary cannot be read at all still does.
    let dictionary: string;
    try {
      const object = parseIndirectObject(
        sourceBytes,
        parsed,
        entry.objectNumber,
        entry.generation,
        readContext,
      );
      const raw = asLatin1(object.bytes);
      const streamBoundary = /\bstream(?:\r\n|\r|\n)/.exec(raw);
      dictionary = streamBoundary ? raw.slice(0, streamBoundary.index) : raw;
    } catch (error) {
      if (!(error instanceof PdfNativeTransactionError)) throw error;
      const dictionaryOnly = readMraiNativeObjectDictionarySource(sourceBytes, {
        objectNumber: entry.objectNumber,
        generation: entry.generation,
      });
      if (dictionaryOnly === null) throw error;
      dictionary = dictionaryOnly;
    }
    for (const match of withoutMarkedContentStreamOwners(dictionary).matchAll(reference)) {
      const key = `${match[1]}:${match[2]}`;
      const count = counts.get(key);
      if (count !== undefined) counts.set(key, count + 1);
    }
  }
  return Object.freeze(targets.map((target) => counts.get(keyFor(target)) ?? 0));
}

export async function inspectMraiNativeContentStream(
  sourceBytes: Uint8Array,
  target: Readonly<{ objectNumber: number; generation: number }>,
  untouchedObjects: readonly Readonly<{
    objectNumber: number;
    generation: number;
    role?: PdfNativeObjectPrecondition["role"];
  }>[] = [],
): Promise<PdfNativeContentStreamInspection> {
  assertObjectReference(target.objectNumber, target.generation);
  const parsed = parsePdf(sourceBytes, false);
  const sourceSha256 = await sha256PdfBytes(sourceBytes);
  return inspectNativeStreamInReadSource(sourceBytes, parsed, sourceSha256, target, untouchedObjects);
}

async function inspectNativeStreamInReadSource(
  sourceBytes: Uint8Array, parsed: ParsedPdf, sourceSha256: string,
  target: Readonly<{ objectNumber: number; generation: number }>,
  untouchedObjects: readonly Readonly<{ objectNumber: number; generation: number; role?: PdfNativeObjectPrecondition["role"] }>[],
  readContext?: NativeObjectReadContext,
): Promise<PdfNativeContentStreamInspection> {
  assertObjectReference(target.objectNumber, target.generation);
  const object = parseIndirectObject(
    sourceBytes,
    parsed,
    target.objectNumber,
    target.generation,
    readContext,
  );
  assertTargetStream(object);
  const untouched = await Promise.all(
    untouchedObjects.map((item) => {
      assertObjectReference(item.objectNumber, item.generation);
      if (
        item.objectNumber === target.objectNumber &&
        item.generation === target.generation
      ) {
        fail("invalid-request", "Target object cannot be an untouched object.");
      }
      return objectPrecondition(
        sourceBytes,
        parsed,
        item.objectNumber,
        item.generation,
        item.role ?? "untouched-object",
        readContext,
      );
    }),
  );
  return Object.freeze({
    identity: Object.freeze({
      sourceSha256,
      objectNumber: target.objectNumber,
      generation: target.generation,
    }),
    objectSha256: await sha256PdfBytes(object.bytes),
    streamSha256: await sha256PdfBytes(object.encodedStreamBytes),
    streamBytes: Uint8Array.from(object.streamBytes),
    untouchedObjects: Object.freeze(untouched),
  });
}

/** Private-byte read index. Its parsed xref is never supplied to a writer.
 * All object/stream inspections retain the ordinary ownership and hash checks. */
export async function createMraiNativeSourceReadInspection(sourceBytes: Uint8Array) {
  if (!sourceBytes.length || sourceBytes.length > MAX_MRAI_NATIVE_SOURCE_BYTES)
    fail("unsupported-document", "PDF source exceeds the byte budget.");
  const bytes = Uint8Array.from(sourceBytes);
  let parsed: ParsedPdf;
  let sourceSha256: string;
  try { parsed = parsePdf(bytes, false); sourceSha256 = await sha256PdfBytes(bytes); }
  catch (error) { bytes.fill(0); throw error; }
  const readContext: NativeObjectReadContext = { source: asLatin1(bytes), members: new Map() };
  const clear = () => {
    bytes.fill(0);
    readContext.source = "";
    for (const members of readContext.members.values())
      for (const member of members) member.bytes.fill(0);
    readContext.members.clear();
  };
  let disposed = false, active = 0;
  const read = async <T>(operation: () => Promise<T>): Promise<T> => {
    if (disposed) fail("source-mismatch", "PDF read source is closed.");
    active++;
    try {
      const result = await operation();
      if (disposed) fail("source-mismatch", "PDF read source was closed.");
      return result;
    } finally { active--; if (disposed && !active) clear(); }
  };
  return {
    sourceSha256,
    inspectObject: (target: Parameters<typeof inspectMraiNativeObject>[1]) =>
      read(() => inspectNativeObjectInReadSource(bytes, parsed, sourceSha256, target, readContext)),
    inspectContentStream: (target: Parameters<typeof inspectMraiNativeContentStream>[1],
      untouchedObjects: Parameters<typeof inspectMraiNativeContentStream>[2] = []) =>
      read(() => inspectNativeStreamInReadSource(bytes, parsed, sourceSha256, target, untouchedObjects, readContext)),
    dispose() { disposed = true; if (!active) clear(); },
  };
}


type NativeContentOperand = Readonly<{
  kind: "name" | "value";
  start: number;
  end: number;
  name?: string;
}>;

interface NativeImageInvocation {
  readonly resourceName: string;
  readonly stream: Readonly<{ objectNumber: number; generation: number }>;
  readonly operatorSpan: Readonly<{ start: number; end: number }>;
  /** The CTM in force at the `Do`, from the `q`/`Q`/`cm` of the drawing it
   * was scanned in (identity at that drawing's start); null when the walk
   * cannot know it (a malformed `cm`, a `Q` with nothing to restore). */
  readonly matrix: PdfNativeEmbeddedImageTransformMatrix | null;
}

/** A Form XObject whose own content stream paints pictures, reached from the
 * page through forms that are each drawn exactly once in the whole document.
 * Writing such a stream changes the one place it is drawn and no other; a
 * form drawn more than once (a letterhead on every page) is not a host, and
 * the pictures inside it stay preserve-only. */
interface NativeImageFormHost {
  /** The form, which is also the content stream a writer rewrites. */
  readonly form: Readonly<{ objectNumber: number; generation: number }>;
  /** Where the form's names resolve: itself, its own Resources object, or --
   * for a form that declares none -- the resources of what draws it. */
  readonly resourceOwner: Readonly<{ objectNumber: number; generation: number }>;
  readonly resourceKind: PdfNativeResourceDictionaryKind;
  readonly images: ReadonlyMap<string, Readonly<{ objectNumber: number; generation: number }>>;
  /** The forms entered from the page to reach this one, outermost first and
   * this one last. */
  readonly chain: readonly Readonly<{ objectNumber: number; generation: number }>[];
  /** The drawing offset of each form `Do` on the way in. */
  readonly orderPrefix: readonly number[];
}

/** One picture paint with the drawing that holds it: null for the page's own
 * streams. `orderKey` places it in the order a reader paints the page: the
 * drawing offset of each form `Do` on the way in, then its own offset. */
interface NativeHostedImageInvocation extends NativeImageInvocation {
  readonly host: NativeImageFormHost | null;
  readonly orderKey: readonly number[];
}

function compareNativeImageOrderKeys(
  left: readonly number[],
  right: readonly number[],
): number {
  for (let index = 0; index < Math.min(left.length, right.length); index += 1) {
    if (left[index] !== right[index]) return left[index] - right[index];
  }
  return left.length - right.length;
}

interface NativeDrawingSegment {
  readonly stream: Readonly<{ objectNumber: number; generation: number }>;
  readonly start: number;
  readonly end: number;
}

const NATIVE_IDENTITY_MATRIX = Object.freeze([
  1, 0, 0, 1, 0, 0,
]) as PdfNativeEmbeddedImageTransformMatrix;

/** A writer aimed by a reader's measure of a picture rewrites only a paint
 * drawn where that picture is: the order a reader counts can name another
 * paint (see `resolvePdfNativeEmbeddedImageOrder`), and the server cannot
 * tell which picture the person meant. */
function assertNativeImagePaintMatrix(
  invocation: NativeImageInvocation,
  expected: readonly number[] | null | undefined,
): void {
  if (expected === undefined || expected === null) return;
  if (!pdfNativeEmbeddedImagePaintMatrixMatches(invocation.matrix, expected)) {
    fail(
      "precondition-mismatch",
      "Embedded image paint is not drawn where the selected picture is.",
    );
  }
}

/**
 * Every `/Name Do` in one page drawing, each named against the stream that
 * holds it. A page drawn from several streams is scanned as the one drawing a
 * reader sees, so an operation that runs across a stream boundary is read
 * whole instead of ending its stream mid-operation.
 *
 * The CTM is followed as a reader follows it (`q` saves it, `Q` restores it,
 * `cm` concatenates onto it), so each paint also says where it draws: what a
 * picture measured on screen is matched against before a writer touches it.
 */
function scanNativeImageInvocations(
  drawingBytes: Uint8Array,
  segments: readonly NativeDrawingSegment[],
): readonly NativeImageInvocation[] {
  const source = asLatin1(drawingBytes);
  const invocations: NativeImageInvocation[] = [];
  let operands: NativeContentOperand[] = [];
  let cursor = 0;
  let tokenCount = 0;
  let ctm: PdfNativeEmbeddedImageTransformMatrix | null = NATIVE_IDENTITY_MATRIX;
  const saved: (PdfNativeEmbeddedImageTransformMatrix | null)[] = [];
  while (cursor < source.length) {
    cursor = skipWhitespace(source, cursor);
    if (cursor >= source.length) break;
    tokenCount += 1;
    if (tokenCount > MAX_CONTENT_STREAM_TOKENS) {
      fail("unsupported-document", "PDF content stream has too many tokens.");
    }
    const start = cursor;
    const character = source[cursor];
    const next = source[cursor + 1];
    if (character === "(") {
      cursor = pdfLiteralEnd(source, cursor, source.length);
      operands.push(Object.freeze({ kind: "value", start, end: cursor }));
      continue;
    }
    if (character === "[") {
      cursor = pdfArrayEnd(source, cursor, source.length);
      operands.push(Object.freeze({ kind: "value", start, end: cursor }));
      continue;
    }
    if (character === "<" && next === "<") {
      cursor = dictionaryEnd(source, cursor);
      operands.push(Object.freeze({ kind: "value", start, end: cursor }));
      continue;
    }
    if (character === "<") {
      cursor = pdfHexEnd(source, cursor, source.length);
      operands.push(Object.freeze({ kind: "value", start, end: cursor }));
      continue;
    }
    if (character === "/") {
      cursor = pdfNameEnd(source, cursor, source.length);
      // Any legal PDF name may stand here: a marked-content tag, a property
      // list name, an intent. `decodePdfName` already bounds its length and
      // proves its `#xx` escapes, and only the name a `Do` consumes has to be
      // a resource key this engine could also write.
      const name = decodePdfName(source, start, cursor);
      operands.push(Object.freeze({ kind: "name", start, end: cursor, name }));
      continue;
    }
    if ("<>]{}/%".includes(character)) {
      fail("unsupported-document", "PDF content stream token is malformed.");
    }
    cursor = pdfTokenEnd(source, cursor, source.length);
    const token = source.slice(start, cursor);
    if (token === "true" || token === "false" || token === "null") {
      operands.push(Object.freeze({ kind: "value", start, end: cursor }));
      continue;
    }
    if (/^[+-.0-9]/.test(token)) {
      if (!/^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(token)) {
        fail("unsupported-document", "PDF content number is malformed.");
      }
      operands.push(Object.freeze({ kind: "value", start, end: cursor }));
      continue;
    }
    if (token === "BI") {
      // One opaque operation: `BI`, the image's own dictionary, `ID`, binary
      // data that may hold anything, `EI`. A reader steps over it whole, and
      // so does this walker -- an inline image carries no resource name and
      // stays unselectable, but the page around it keeps its XObject paints.
      cursor = skipNativeInlineImage(source, cursor);
      operands = [];
      continue;
    }
    if (token === "q") {
      saved.push(ctm);
    } else if (token === "Q") {
      // A restore with nothing saved: readers disagree on it (one keeps the
      // state, another resets it), so what follows has no known place.
      ctm = saved.length > 0 ? (saved.pop() ?? null) : null;
    } else if (token === "cm") {
      const values =
        operands.length === 6 && operands.every((operand) => operand.kind === "value")
          ? operands.map((operand) => Number(source.slice(operand.start, operand.end)))
          : null;
      ctm =
        ctm && values && values.every((value) => Number.isFinite(value))
          ? concatPdfNativeEmbeddedImageTransformMatrices(
              values as unknown as PdfNativeEmbeddedImageTransformMatrix,
              ctm,
            )
          : null;
    }
    if (token === "Do") {
      if (operands.length !== 1 || operands[0].kind !== "name") {
        fail("unsupported-document", "PDF XObject invocation is ambiguous.");
      }
      if (!PDF_RESOURCE_NAME.test(operands[0].name!)) {
        fail(
          "unsupported-document",
          "PDF content resource name is unsupported.",
        );
      }
      const owner = segments.find(
        (segment) =>
          operands[0].start >= segment.start && cursor <= segment.end,
      );
      if (!owner) {
        fail(
          "unsupported-document",
          "PDF XObject invocation crosses a page content stream boundary.",
        );
      }
      invocations.push(
        Object.freeze({
          resourceName: operands[0].name!,
          stream: Object.freeze({ ...owner.stream }),
          operatorSpan: Object.freeze({
            start: operands[0].start - owner.start,
            end: cursor - owner.start,
          }),
          matrix: ctm,
        }),
      );
    }
    operands = [];
  }
  // The producer may divide the drawing at any token boundary -- Acrobat
  // routinely ends a piece with the operands whose operator opens the next
  // one -- but the walk above is over the whole drawing, so anything still
  // pending here is past its last operator. A reader drops it, and so does
  // this walker; the one exception is a single pending name, which a piece
  // this page does not draw could still turn into a `Do`, and this walker has
  // to list every paint on the page in order.
  if (operands.length === 1 && operands[0].kind === "name") {
    fail(
      "unsupported-document",
      "PDF content stream ends inside an operation.",
    );
  }
  return Object.freeze(invocations);
}

/** Step over one inline image, from just after its `BI` to just after its
 * `EI`. Its dictionary is read only to learn how many bytes of data it wrote:
 * the data may contain `EI` itself, which is why a stated count comes first. */
export function skipMraiNativeInlineImage(source: string, afterBegin: number): number {
  return skipNativeInlineImage(source, afterBegin);
}

function skipNativeInlineImage(source: string, afterBegin: number): number {
  let cursor = afterBegin;
  while (cursor < source.length) {
    cursor = skipWhitespace(source, cursor);
    if (cursor >= source.length) break;
    const character = source[cursor];
    const next = source[cursor + 1];
    if (character === "(") {
      cursor = pdfLiteralEnd(source, cursor, source.length);
      continue;
    }
    if (character === "[") {
      cursor = pdfArrayEnd(source, cursor, source.length);
      continue;
    }
    if (character === "<" && next === "<") {
      cursor = dictionaryEnd(source, cursor);
      continue;
    }
    if (character === "<") {
      cursor = pdfHexEnd(source, cursor, source.length);
      continue;
    }
    if (character === "/") {
      cursor = pdfNameEnd(source, cursor, source.length);
      continue;
    }
    if ("<>]{}/%".includes(character)) {
      fail("unsupported-document", "PDF inline image dictionary is malformed.");
    }
    const tokenEnd = pdfTokenEnd(source, cursor, source.length);
    if (source.slice(cursor, tokenEnd) !== "ID") {
      cursor = tokenEnd;
      continue;
    }
    // Exactly one whitespace byte separates `ID` from the data it introduces.
    const dataEnd = mraiNativeInlineImageDataEnd(
      source,
      tokenEnd + 1,
      source.slice(afterBegin, cursor),
    );
    if (dataEnd === null) {
      fail("unsupported-document", "Inline PDF image data is unterminated.");
    }
    return dataEnd;
  }
  fail("unsupported-document", "Inline PDF image has no data.");
}

function exactNamedObjectType(
  objectBytes: Uint8Array,
  key: "Type" | "Subtype",
  expected: string,
  label: string,
  /** Read a key the specification marks optional as satisfied when it is
   * absent. An XObject's `/Type` is optional (PDF 32000-1 Table 93 for an
   * image, Table 95 for a form): its `/Subtype` is what says what it is, and
   * plenty of producers leave `/Type` out. Absence is not ambiguity -- a key
   * that IS written still has to be written once and say the expected name.
   * Mirror of the text route's form walk, which already reads it this way. */
  optional = false,
): void {
  const entries = parseMraiNativeTopLevelDictionaryNameEntries(objectBytes);
  const matches = entries.filter((entry) => entry.key === key);
  if (optional && matches.length === 0) return;
  if (matches.length !== 1 || matches[0].directName !== expected) {
    fail("unsupported-document", `${label} ${key} is missing or ambiguous.`);
  }
}

function sameNativeObjectRef(
  left: Readonly<{ objectNumber: number; generation: number }>,
  right: Readonly<{ objectNumber: number; generation: number }>,
): boolean {
  return (
    left.objectNumber === right.objectNumber &&
    left.generation === right.generation
  );
}

/**
 * What each XObject name of one page's resources is (ISO 32000-1 Table 92):
 * an Image, a Form or a PostScript XObject. A layer move of the page's
 * content (vector-content-transaction: the Selection pane's objects) tells
 * a picture's `Do` from a Form's by it. A name whose XObject says no exact
 * subtype is left out: the caller treats a `Do` it cannot name as content
 * it does not address.
 */
export async function readMraiNativePageXObjectKinds(
  sourceBytes: Uint8Array,
  pageNumber: number,
): Promise<ReadonlyMap<string, "Image" | "Form" | "TextForm" | "PS">> {
  const pages = await inspectMraiNativeDocumentPageContentGraphs(sourceBytes);
  const page = pages[pageNumber - 1];
  if (!page) fail("invalid-request", "Selected PDF page is out of range.");
  const resourceOwner = await inspectMraiNativeObject(sourceBytes, page.resourceAuthority.owner);
  const mappings = xObjectDictionaryEntries(
    sourceBytes,
    resourceOwner.objectBytes,
    page.resourceAuthority.dictionaryKind,
    `PDF page ${pageNumber}`,
    true,
  );
  const kinds = new Map<string, "Image" | "Form" | "TextForm" | "PS">();
  for (const mapping of mappings) {
    const dictionary = readMraiNativeObjectDictionarySource(sourceBytes, mapping.target);
    if (dictionary === null) continue;
    const subtype = parseMraiNativeTopLevelDictionaryNameEntries(latin1Bytes(dictionary)).filter(
      (entry) => entry.key === "Subtype",
    );
    const name = subtype.length === 1 ? (subtype[0]?.directName ?? null) : null;
    if (name === "Form") {
      // A Form shows text with the text state it is drawn in when its own
      // drawing (or a Form it draws) shows any; one this reader cannot read
      // is taken to.
      let drawing: string | null = null;
      try {
        const inspected = await inspectMraiNativeObject(sourceBytes, mapping.target);
        drawing = inspected.streamBytes ? asLatin1(inspected.streamBytes) : null;
      } catch {
        drawing = null;
      }
      kinds.set(mapping.resourceName, drawing === null || /(^|[^A-Za-z])(BT|Do)(?![A-Za-z])/.test(drawing)
        ? "TextForm" : "Form");
    } else if (name === "Image" || name === "PS") {
      kinds.set(mapping.resourceName, name);
    }
  }
  return kinds;
}

async function resolveNativeImageRemovalGraph(
  sourceBytes: Uint8Array,
  pageNumber: number,
): Promise<
  Readonly<{
    page: PdfNativePageContentGraphInspection;
    resourceOwner: PdfNativeObjectInspection;
    images: ReadonlyMap<
      string,
      Readonly<{ objectNumber: number; generation: number }>
    >;
    invocations: readonly NativeHostedImageInvocation[];
    formHosts: readonly NativeImageFormHost[];
    segments: readonly NativeDrawingSegment[];
    contentStreamReferenceCounts: ReadonlyMap<string, number>;
  }>
> {
  const pages = await inspectMraiNativeDocumentPageContentGraphs(sourceBytes);
  const page = pages[pageNumber - 1];
  if (!page) {
    fail("invalid-request", "Selected PDF page is out of range.");
  }
  const contentStreamReferenceCounts = new Map<string, number>();
  for (const candidatePage of pages) {
    for (const contentStream of candidatePage.contentStreams) {
      const key = `${contentStream.objectNumber}:${contentStream.generation}`;
      contentStreamReferenceCounts.set(
        key,
        (contentStreamReferenceCounts.get(key) ?? 0) + 1,
      );
    }
  }
  const resourceOwner = await inspectMraiNativeObject(
    sourceBytes,
    page.resourceAuthority.owner,
  );
  // A page whose resources declare no /XObject paints no pictures; that is
  // no pictures, not a malformed dictionary (mirror of replay's read-through).
  const mappings = xObjectDictionaryEntries(
    sourceBytes,
    resourceOwner.objectBytes,
    page.resourceAuthority.dictionaryKind,
    `PDF page ${pageNumber}`,
    true,
  );
  const images = new Map<
    string,
    Readonly<{ objectNumber: number; generation: number }>
  >();
  // Names the page maps to an XObject that is not a picture. Wrapping a logo,
  // an overprint or a chart in a Form XObject is ordinary (headers, footers,
  // forms, reports all do it), and a transparency group is written that way
  // too, so most pages that paint pictures also map a form beside them. Such
  // an entry is not a picture and not a defect: it simply is not a paint this
  // route addresses, exactly like a font or a pattern the page also maps.
  // Refusing the page over it took every picture ON the page down with it.
  //
  // A picture drawn INSIDE such a form stays preserve-only, and is a known
  // boundary rather than an oversight: its `Do` is in the form's own content
  // stream against the form's own `/Resources`, so addressing it means
  // writing that stream, and a form the document references more than once
  // draws in more than one place -- editing it would silently change every
  // one. Supporting it needs the text route's shape (`resolveFormPaths` and
  // `formInstanceIsolation`), which counts instances and clones the chain
  // first; until then the page lists the paints it can address and no more.
  const nonImages = new Set<string>();
  const forms = new Map<string, Readonly<{ objectNumber: number; generation: number }>>();
  for (const mapping of mappings) {
    // Moving or removing a paint rewrites the page, never the picture: its
    // dictionary says what it is, and a picture whose stream this engine
    // cannot decode (JPX, CCITT, a predictor it does not read) is still a
    // picture that can be taken off the page or put elsewhere on it.
    const targetDictionary = readMraiNativeObjectDictionarySource(
      sourceBytes,
      mapping.target,
    );
    if (targetDictionary === null) {
      fail("unsupported-document", "PDF XObject is not readable.");
    }
    const targetBytes = latin1Bytes(targetDictionary);
    exactNamedObjectType(targetBytes, "Type", "XObject", "PDF XObject", true);
    const subtype = parseMraiNativeTopLevelDictionaryNameEntries(
      targetBytes,
    ).filter((entry) => entry.key === "Subtype");
    if (subtype.length !== 1 || subtype[0].directName === null) {
      fail("unsupported-document", "PDF XObject Subtype is ambiguous.");
    }
    // PDF 32000-1 Table 92 names three XObject subtypes. A form paints its
    // own content stream and a PostScript XObject is ignored by a conforming
    // reader (7.8.4.2); neither is a picture this route can address, and a
    // subtype outside that list is a structure this engine has never read.
    if (
      subtype[0].directName === "Form" ||
      subtype[0].directName === "PS"
    ) {
      nonImages.add(mapping.resourceName);
      if (subtype[0].directName === "Form") {
        forms.set(mapping.resourceName, mapping.target);
      }
      continue;
    }
    if (subtype[0].directName !== "Image") {
      fail("unsupported-document", "Unsupported PDF XObject subtype.");
    }
    images.set(mapping.resourceName, mapping.target);
  }
  // The page's streams are one drawing, so the paints are scanned across the
  // whole of it and only then attributed to the stream that holds each one.
  const contentStreamInspections = await inspectMraiNativePageContentStreams(
    sourceBytes,
    pageNumber,
  );
  const drawingBytes = joinNativePageDrawing(contentStreamInspections);
  const segments = contentStreamInspections.map((inspection) => ({
    stream: inspection.contentStream,
    start: inspection.drawingOffset,
    end: inspection.drawingOffset + inspection.streamBytes.length,
  }));
  const scanned = scanNativeImageInvocations(drawingBytes, segments);
  // The page's names are its own only when no other page reads them: its
  // resources are written in the page itself or in an object nothing else
  // references.
  const pageOwnerExclusive =
    sameNativeObjectRef(page.resourceAuthority.owner, page.page) ||
    countMraiNativeIndirectObjectReferences(
      sourceBytes,
      page.resourceAuthority.owner,
    ) === 1;
  // A form mapped through resources other pages share is reachable from each
  // of them, but it is drawn only where a content stream invokes its name.
  // One such dictionary for the whole document is common (each page draws
  // its own form from it), so the other pages' drawings decide, read once.
  let namesDrawnElsewhere: Promise<ReadonlySet<string> | null> | null = null;
  const drawnByAnotherPage = async (resourceName: string) => {
    if (pageOwnerExclusive) return false;
    namesDrawnElsewhere ??= resourceNamesDrawnByOtherPages({
      sourceBytes,
      pages,
      pageNumber,
      owner: page.resourceAuthority.owner,
    });
    const names = await namesDrawnElsewhere;
    return names === null || names.has(resourceName);
  };
  const invocations: NativeHostedImageInvocation[] = [];
  const formHosts: NativeImageFormHost[] = [];
  for (const invocation of scanned) {
    const segment = segments.find((item) =>
      sameNativeObjectRef(item.stream, invocation.stream),
    )!;
    const drawingOffset = segment.start + invocation.operatorSpan.start;
    const form = forms.get(invocation.resourceName);
    if (form) {
      // A form's own pictures are painted where the form is, so they take
      // their place in the page's order at its `Do`.
      const hosted = await nativeImageFormHostInvocations({
        sourceBytes,
        form,
        parentOwner: page.resourceAuthority.owner,
        parentKind: page.resourceAuthority.dictionaryKind,
        parentOwnerBytes: resourceOwner.objectBytes,
        parentOwnerExclusive:
          !(await drawnByAnotherPage(invocation.resourceName)) &&
          contentStreamReferenceCounts.get(
            `${invocation.stream.objectNumber}:${invocation.stream.generation}`,
          ) === 1,
        invocationCount: scanned.filter(
          (item) => item.resourceName === invocation.resourceName,
        ).length,
        chain: Object.freeze([]),
        orderPrefix: Object.freeze([drawingOffset]),
        parentMatrix: invocation.matrix,
      });
      if (hosted) {
        invocations.push(...hosted.invocations);
        formHosts.push(...hosted.hosts);
      }
      continue;
    }
    // A `Do` of a PostScript XObject is a paint this route leaves alone; it
    // is stepped over so the pictures around it keep their order.
    if (nonImages.has(invocation.resourceName)) continue;
    if (!images.has(invocation.resourceName)) {
      fail(
        "unsupported-document",
        "PDF XObject invocation has no exact image resource.",
      );
    }
    invocations.push(
      Object.freeze({
        ...invocation,
        host: null,
        orderKey: Object.freeze([drawingOffset]),
      }),
    );
  }
  return Object.freeze({
    page,
    resourceOwner,
    images,
    invocations: Object.freeze(invocations),
    formHosts: Object.freeze(formHosts),
    segments: Object.freeze(segments.map((segment) => Object.freeze({ ...segment }))),
    contentStreamReferenceCounts,
  });
}

type NativeImageGraph = Awaited<ReturnType<typeof resolveNativeImageRemovalGraph>>;

/** The picture a paint names, resolved where its drawing resolves names. */
function nativeImageOf(
  graph: NativeImageGraph,
  invocation: NativeHostedImageInvocation,
): Readonly<{ objectNumber: number; generation: number }> | undefined {
  return (invocation.host ? invocation.host.images : graph.images).get(
    invocation.resourceName,
  );
}

/** The object whose resource dictionary names the paint's picture. */
function nativeImageResourceOwner(
  graph: NativeImageGraph,
  host: NativeImageFormHost | null,
): Readonly<{ objectNumber: number; generation: number }> {
  return host ? host.resourceOwner : graph.page.resourceAuthority.owner;
}

/** A page stream a writer rewrites must be drawn by this page alone; a form
 * host was proven drawn once when the graph admitted it. */
function assertNativeImagePaintStream(
  graph: NativeImageGraph,
  invocation: NativeHostedImageInvocation,
): void {
  if (invocation.host) return;
  const key = `${invocation.stream.objectNumber}:${invocation.stream.generation}`;
  if (graph.contentStreamReferenceCounts.get(key) !== 1) {
    fail(
      "unsupported-document",
      "Embedded image content stream is shared across page Contents.",
    );
  }
}

/** Every object a paint's writer leaves exactly as it is: the page, where its
 * names resolve, the picture, the page's other streams and every form on the
 * way to the drawing it rewrites. */
function nativeImagePaintPins(
  graph: NativeImageGraph,
  invocation: NativeHostedImageInvocation,
  image: Readonly<{ objectNumber: number; generation: number }>,
): readonly Readonly<{ objectNumber: number; generation: number }>[] {
  return [
    graph.page.page,
    graph.page.resourceAuthority.owner,
    ...(invocation.host
      ? [invocation.host.resourceOwner, ...invocation.host.chain]
      : []),
    image,
    ...graph.page.contentStreams,
  ].filter(
    (item, index, items) =>
      !sameNativeObjectRef(item, invocation.stream) &&
      items.findIndex((candidate) => sameNativeObjectRef(candidate, item)) ===
        index,
  );
}

/** The drawing-order key of a position in a page stream or a form host. */
function nativeImageOrderKeyAt(
  graph: NativeImageGraph,
  stream: Readonly<{ objectNumber: number; generation: number }>,
  offset: number,
): readonly number[] | null {
  const host = graph.formHosts.find((item) => sameNativeObjectRef(item.form, stream));
  if (host) return Object.freeze([...host.orderPrefix, offset]);
  const segment = graph.segments.find((item) => sameNativeObjectRef(item.stream, stream));
  return segment ? Object.freeze([segment.start + offset]) : null;
}

/** Where a form's names resolve, read from its own dictionary: inline, as a
 * Resources object of their own, or not at all (it then reads the resources
 * of what draws it). Null when the entry is ambiguous or unreadable. */
function nativeFormResourceAuthority(
  dictionary: string,
):
  | Readonly<{ kind: "inherit" }>
  | Readonly<{ kind: "nested-resources" }>
  | Readonly<{
      kind: "resource-dictionary";
      ref: Readonly<{ objectNumber: number; generation: number }>;
    }>
  | null {
  const entries = topLevelDictionaryEntries(dictionary).filter(
    (entry) => entry.key === "Resources",
  );
  if (entries.length === 0) return Object.freeze({ kind: "inherit" as const });
  if (entries.length !== 1) return null;
  const value = dictionary.slice(entries[0].valueStart, entries[0].valueEnd).trim();
  if (value.startsWith("<<")) {
    return Object.freeze({ kind: "nested-resources" as const });
  }
  const reference = /^(\d+)\s+(\d+)\s+R$/.exec(value);
  return reference
    ? Object.freeze({
        kind: "resource-dictionary" as const,
        ref: Object.freeze({
          objectNumber: Number(reference[1]),
          generation: Number(reference[2]),
        }),
      })
    : null;
}

const MAX_NATIVE_IMAGE_FORM_DEPTH = 8;

/**
 * The pictures one form paints, and those of the forms it draws in turn, when
 * the form is drawn exactly once in the whole document: its object has one
 * reference, the drawing that invokes it names it once, and that drawing's
 * names belong to it alone. Null when any of that fails or the form cannot
 * be read -- such a form keeps its pictures preserve-only, exactly as every
 * form did before, and the page's own pictures are unaffected.
 */
/**
 * A Form XObject the page draws directly, exactly once, that nothing else in
 * the document references: the one place its content is drawn, so a vector
 * inside it can be edited by rewriting the form's own stream. Its `Do`'s
 * offset in the page drawing gives the graphics state it inherits; its
 * `/Matrix` and `/BBox` are read as the server reads them (a top-level
 * Matrix entry, exactly one BBox). Mirror of the server's form target in
 * _vector_page_authority.
 */
/** One Form on the way to a host: where its parent draws it, and the
 * transform and clip it adds. */
export interface PdfNativeVectorFormHostLevel {
  readonly resourceName: string;
  readonly form: Readonly<{ objectNumber: number; generation: number }>;
  /** Offset of the `Do` operand: in the page drawing (the joined streams)
   * for the first Form, in the parent Form's decoded stream otherwise. */
  readonly doOffset: number;
  readonly matrix: readonly [number, number, number, number, number, number];
  readonly bbox: readonly [number, number, number, number] | null;
  readonly streamInspection: PdfNativeContentStreamInspection;
}

export interface PdfNativeVectorFormHost {
  readonly resourceName: string;
  readonly form: Readonly<{ objectNumber: number; generation: number }>;
  /** Offset of the first `Do` on the way in the page drawing (the joined
   * streams): the host's own for a Form the page draws. */
  readonly doDrawingOffset: number;
  readonly matrix: readonly [number, number, number, number, number, number];
  readonly bbox: readonly [number, number, number, number] | null;
  readonly streamInspection: PdfNativeContentStreamInspection;
  /** The Forms entered to reach this one, outermost first, ending with it:
   * one level for a Form the page draws, more for one drawn inside another. */
  readonly path: readonly PdfNativeVectorFormHostLevel[];
  /** Present when other pages reference the form too (a logo in every
   * header), when the page draws it more than once, or when it is drawn
   * inside another Form: an edit here is written on a clone the page's names
   * then reach, so every other drawing keeps the original -- the
   * clone-on-write an in-form text edit has always used (mirror of the
   * server's form isolation). */
  readonly formInstanceIsolation?: PdfNativeFormInstanceIsolationPlan;
  /** Which drawing of its Form this host is, when the page shows the Form
   * more than once: the `Do` offsets on its way, outermost first. Proofs read
   * through it carry it, so every drawing's vectors are told apart. */
  readonly formInstance?: string;
}

/** How deep the vector engine follows Forms drawn inside Forms. */
const MAX_VECTOR_FORM_DEPTH = 3;

const FORM_NUMBER = "([-+]?(?:\\d+(?:\\.\\d*)?|\\.\\d+))";
const FORM_MATRIX_ARRAY = new RegExp(
  `^\\[\\s*${FORM_NUMBER}\\s+${FORM_NUMBER}\\s+${FORM_NUMBER}\\s+${FORM_NUMBER}\\s+${FORM_NUMBER}\\s+${FORM_NUMBER}\\s*\\]$`,
);
const FORM_BBOX_ARRAY = new RegExp(
  `/BBox\\s*\\[\\s*${FORM_NUMBER}\\s+${FORM_NUMBER}\\s+${FORM_NUMBER}\\s+${FORM_NUMBER}\\s*\\]`,
  "g",
);

/** A form's own `/Matrix` (a top-level entry; identity when absent), null
 * when malformed or ambiguous (mirror of the server's _form_matrix_from_dictionary). */
function formMatrixFromDictionary(
  dictionary: string,
): readonly [number, number, number, number, number, number] | null {
  let entries: readonly TopLevelDictionaryEntry[];
  try {
    entries = topLevelDictionaryEntries(dictionary);
  } catch (error) {
    if (error instanceof PdfNativeTransactionError) {
      return dictionary.includes("/Matrix") ? null : [1, 0, 0, 1, 0, 0];
    }
    throw error;
  }
  const entry = entries.find((item) => item.key === "Matrix");
  if (!entry) return [1, 0, 0, 1, 0, 0];
  const matched = FORM_MATRIX_ARRAY.exec(dictionary.slice(entry.valueStart, entry.valueEnd).trim());
  if (!matched) return null;
  const values = matched.slice(1, 7).map(Number);
  if (values.some((value) => !Number.isFinite(value) || Math.abs(value) > 1_000_000)) return null;
  return [values[0], values[1], values[2], values[3], values[4], values[5]];
}

/** A form's `/BBox`, when it is written exactly once (mirror of the server's
 * _form_footprint_from_dictionary reading). */
function formBBoxFromDictionary(dictionary: string): readonly [number, number, number, number] | null {
  const boxes = [...dictionary.matchAll(FORM_BBOX_ARRAY)];
  if (boxes.length !== 1) return null;
  const values = boxes[0].slice(1, 5).map(Number);
  if (values.some((value) => !Number.isFinite(value))) return null;
  return [values[0], values[1], values[2], values[3]];
}

function formOwnsResources(sourceBytes: Uint8Array, form: Readonly<{ objectNumber: number; generation: number }>): boolean {
  const dictionary = readMraiNativeObjectDictionarySource(sourceBytes, form);
  if (dictionary === null) return false;
  try {
    return topLevelDictionaryEntries(dictionary).some((entry) => entry.key === "Resources");
  } catch (error) {
    if (error instanceof PdfNativeTransactionError) return false;
    throw error;
  }
}

/** Whether a form of the page could draw another of the page's forms
 * through the page's names: it declares no resources of its own and its
 * content invokes an XObject at all. PowerPoint writes pictures as forms
 * without resources that only paint paths; such a form draws no other form,
 * and the page's drawing alone still counts every instance. An unreadable
 * form may draw anything (mirror of the server's _vector_form_host). */
async function formMayDrawThroughPageNames(
  sourceBytes: Uint8Array,
  form: Readonly<{ objectNumber: number; generation: number }>,
): Promise<boolean> {
  if (formOwnsResources(sourceBytes, form)) return false;
  try {
    const inspection = await inspectMraiNativeContentStream(sourceBytes, form);
    return (
      scanNativeImageInvocations(inspection.streamBytes, [
        { stream: form, start: 0, end: inspection.streamBytes.length },
      ]).length > 0
    );
  } catch (error) {
    if (error instanceof PdfNativeTransactionError) return true;
    throw error;
  }
}

export async function inspectMraiNativeVectorFormHosts(
  sourceBytes: Uint8Array,
  pageNumber: number,
): Promise<readonly PdfNativeVectorFormHost[]> {
  const pages = await inspectMraiNativeDocumentPageContentGraphs(sourceBytes);
  const page = pages[pageNumber - 1];
  if (!page) return Object.freeze([]);
  const contentStreamReferenceCounts = new Map<string, number>();
  for (const candidatePage of pages) {
    for (const contentStream of candidatePage.contentStreams) {
      const key = `${contentStream.objectNumber}:${contentStream.generation}`;
      contentStreamReferenceCounts.set(key, (contentStreamReferenceCounts.get(key) ?? 0) + 1);
    }
  }
  const resourceOwner = await inspectMraiNativeObject(sourceBytes, page.resourceAuthority.owner);
  let mappings: readonly Readonly<{ resourceName: string; target: Readonly<{ objectNumber: number; generation: number }> }>[];
  try {
    mappings = xObjectDictionaryEntries(
      sourceBytes,
      resourceOwner.objectBytes,
      page.resourceAuthority.dictionaryKind,
      `PDF page ${pageNumber}`,
      true,
    );
  } catch (error) {
    if (error instanceof PdfNativeTransactionError) return Object.freeze([]);
    throw error;
  }
  const forms = new Map<string, Readonly<{ objectNumber: number; generation: number }>>();
  // A form of this page that declares no resources of its own reads the
  // page's, so it could draw another form too; the page drawing alone cannot
  // then count that form's instances (mirror of the server's _vector_form_host).
  let everyFormOwnsResources = true;
  for (const mapping of mappings) {
    const dictionary = readMraiNativeObjectDictionarySource(sourceBytes, mapping.target);
    if (dictionary === null) continue;
    const names = parseMraiNativeTopLevelDictionaryNameEntries(latin1Bytes(dictionary)).filter(
      (entry) => entry.key === "Subtype",
    );
    if (names.length !== 1 || names[0].directName !== "Form") continue;
    // The drawing of a table, chart or SmartArt made in MRAI, which the
    // page draws as its own content (managed_object_content.py names the
    // object on the form): its rules and bars are the object's, drawn again
    // from its model at the next save, and no shapes of the page to move or
    // restyle one by one. (Such a form has resources of its own.)
    let managedObject = /\/MRNoteObject(?![A-Za-z0-9])/.test(dictionary);
    if (managedObject) {
      try {
        managedObject = topLevelDictionaryEntries(dictionary).some((entry) => entry.key === "MRNoteObject");
      } catch (error) {
        if (!(error instanceof PdfNativeTransactionError)) throw error;
      }
    }
    if (managedObject) continue;
    forms.set(mapping.resourceName, mapping.target);
    try {
      if (!topLevelDictionaryEntries(dictionary).some((entry) => entry.key === "Resources")) {
        everyFormOwnsResources = false;
      }
    } catch (error) {
      if (!(error instanceof PdfNativeTransactionError)) throw error;
      everyFormOwnsResources = false;
    }
  }
  if (forms.size === 0) return Object.freeze([]);
  const formNameCounts = new Map<string, number>();
  for (const form of forms.values()) {
    const key = `${form.objectNumber}:${form.generation}`;
    formNameCounts.set(key, (formNameCounts.get(key) ?? 0) + 1);
  }
  const contentStreamInspections = await inspectMraiNativePageContentStreams(sourceBytes, pageNumber);
  const drawingBytes = joinNativePageDrawing(contentStreamInspections);
  const segments = contentStreamInspections.map((inspection) => ({
    stream: inspection.contentStream,
    start: inspection.drawingOffset,
    end: inspection.drawingOffset + inspection.streamBytes.length,
  }));
  const scanned = scanNativeImageInvocations(drawingBytes, segments);
  // The page's names are its own only when no other page reads them.
  const ownerIsPage = sameNativeObjectRef(page.resourceAuthority.owner, page.page);
  const pageOwnerExclusive =
    ownerIsPage ||
    countMraiNativeIndirectObjectReferences(sourceBytes, page.resourceAuthority.owner) === 1;
  // A form other pages reference too is written on a clone the page's names
  // then reach: the page's own dictionary is rewritten, or a resources object
  // of its own is cloned first (never names a parent Pages node lends it).
  const resourcesLink = !ownerIsPage && page.resourceAuthority.dictionaryKind === "resource-dictionary";
  const isolable = ownerIsPage || resourcesLink;
  // Forms of the page that could draw another form through the page's names.
  const drawingThroughPageNames: Readonly<{ objectNumber: number; generation: number }>[] = [];
  if (!everyFormOwnsResources) {
    for (const other of forms.values()) {
      if (await formMayDrawThroughPageNames(sourceBytes, other)) drawingThroughPageNames.push(other);
    }
  }
  const parsed = parseSourceStructure(sourceBytes);
  // The clone-on-write rewrites the page and copies the chain: every Form
  // must be a stream object of its own, the page (and a resources object of
  // its own) a dictionary, which the writer writes out as a plain object when
  // the source keeps it in an object stream (mirror of the isolation
  // writer's gate).
  const writableReadings = new Map<string, boolean>();
  const writable = (ref: Readonly<{ objectNumber: number; generation: number }>, stream: boolean) => {
    const key = `${ref.objectNumber}:${ref.generation}:${stream}`;
    const known = writableReadings.get(key);
    if (known !== undefined) return known;
    let result: boolean;
    try {
      const object = parseIndirectObject(sourceBytes, parsed, ref.objectNumber, ref.generation);
      result =
        (object.storage === "direct" || (!stream && object.storage === "object-stream")) &&
        (object.streamBytes !== null) === stream;
    } catch (error) {
      if (!(error instanceof PdfNativeTransactionError)) throw error;
      result = false;
    }
    writableReadings.set(key, result);
    return result;
  };
  const pageWritable =
    isolable &&
    writable(page.page, false) &&
    (!resourcesLink || writable(page.resourceAuthority.owner, false));
  let pagePin: PdfNativeFormInstanceIsolationPlan["page"] | null = null;
  const pageIsolationPin = async () =>
    (pagePin ??= Object.freeze({
      ...page.page,
      expectedObjectSha256: (await inspectMraiNativeObject(sourceBytes, page.page)).objectSha256,
    }));
  const resourcesLinkEdges = resourcesLink
    ? [
        Object.freeze({
          resourceName: "Resources",
          kind: "resources" as const,
          objectNumber: page.resourceAuthority.owner.objectNumber,
          generation: page.resourceAuthority.owner.generation,
          expectedObjectSha256: resourceOwner.objectSha256,
        }),
      ]
    : [];
  // A drawing renamed on the page takes a new name in the page's own direct
  // /XObject dictionary (or the shared resources dictionary cloned for it).
  let pageNames: ReadonlySet<string> | null | undefined;
  const pageDirectNames = () => {
    if (pageNames !== undefined) return pageNames;
    try {
      pageNames = new Set(
        parseDirectXObjectDictionaryEntries(
          resourceOwner.objectBytes,
          page.resourceAuthority.dictionaryKind,
          `PDF page ${pageNumber}`,
        ).map((entry) => entry.resourceName),
      );
    } catch (error) {
      if (!(error instanceof PdfNativeTransactionError)) throw error;
      pageNames = null;
    }
    return pageNames;
  };
  const freshName = (taken: ReadonlySet<string>, operandStart: number) => {
    for (let suffix = 0; suffix < 64; suffix += 1) {
      const name = suffix === 0 ? `MRAIi${operandStart}` : `MRAIi${operandStart}_${suffix}`;
      if (!taken.has(name)) return name;
    }
    return null;
  };
  const invocationCounts = new Map<string, number>();
  for (const invocation of scanned) {
    invocationCounts.set(invocation.resourceName, (invocationCounts.get(invocation.resourceName) ?? 0) + 1);
  }
  // Each Form's references counted once, all in one pass over the source: a
  // chart page draws the same few Forms dozens of times.
  const formList = [...new Map([...forms.values()].map((form) => [`${form.objectNumber}:${form.generation}`, form] as const)).values()];
  const formReferenceCounts = new Map(
    countMraiNativeIndirectObjectReferencesBatch(sourceBytes, formList).map(
      (count, index) => [`${formList[index].objectNumber}:${formList[index].generation}`, count] as const,
    ),
  );
  const hosts: PdfNativeVectorFormHost[] = [];
  for (const invocation of scanned) {
    const form = forms.get(invocation.resourceName);
    if (!form) continue;
    const segment = segments.find((item) => sameNativeObjectRef(item.stream, invocation.stream));
    if (!segment) continue;
    const shared = formReferenceCounts.get(`${form.objectNumber}:${form.generation}`) !== 1;
    // A Form the page draws more than once is edited one drawing at a time:
    // that `Do` is renamed to reach a clone (formInstanceIsolation).
    const instanced = invocationCounts.get(invocation.resourceName)! > 1;
    const isolated = shared || instanced;
    if (
      formNameCounts.get(`${form.objectNumber}:${form.generation}`) !== 1 ||
      contentStreamReferenceCounts.get(`${invocation.stream.objectNumber}:${invocation.stream.generation}`) !== 1 ||
      (isolated ? !pageWritable : !pageOwnerExclusive) ||
      // Isolation re-points the form's name in a cloned resources object; a
      // form that reads those names itself is not reached through one
      // (mirror of the server's resolved form paths).
      (isolated && resourcesLink && !formOwnsResources(sourceBytes, form)) ||
      // Another form of the page without resources of its own could draw
      // this one from the page's names.
      drawingThroughPageNames.some((other) => !sameNativeObjectRef(other, form))
    ) {
      continue;
    }
    const dictionary = readMraiNativeObjectDictionarySource(sourceBytes, form);
    if (dictionary === null) continue;
    const matrix = formMatrixFromDictionary(dictionary);
    if (matrix === null) continue;
    let streamInspection: PdfNativeContentStreamInspection;
    try {
      streamInspection = await inspectMraiNativeContentStream(sourceBytes, form);
    } catch (error) {
      if (error instanceof PdfNativeTransactionError) continue;
      throw error;
    }
    let formInstanceIsolation: PdfNativeFormInstanceIsolationPlan | undefined;
    if (isolated) {
      if (!writable(form, true)) continue;
      let instance: PdfNativeFormInstanceIsolationInstance | undefined;
      if (instanced) {
        const taken = pageDirectNames();
        const resourceName = taken ? freshName(taken, invocation.operatorSpan.start) : null;
        if (!resourceName) continue;
        instance = Object.freeze({
          contentStream: Object.freeze({ ...invocation.stream }),
          operandStart: invocation.operatorSpan.start,
          resourceName,
        });
      }
      formInstanceIsolation = Object.freeze({
        page: await pageIsolationPin(),
        chain: Object.freeze([
          ...resourcesLinkEdges,
          Object.freeze({
            resourceName: invocation.resourceName,
            kind: "xobject" as const,
            objectNumber: form.objectNumber,
            generation: form.generation,
            expectedObjectSha256: streamInspection.objectSha256,
            ...(instance ? { instance } : {}),
          }),
        ]),
      });
    }
    const level: PdfNativeVectorFormHostLevel = Object.freeze({
      resourceName: invocation.resourceName,
      form: Object.freeze({ ...form }),
      doOffset: segment.start + invocation.operatorSpan.start,
      matrix,
      bbox: formBBoxFromDictionary(dictionary),
      streamInspection,
    });
    hosts.push(
      Object.freeze({
        resourceName: level.resourceName,
        form: level.form,
        doDrawingOffset: level.doOffset,
        matrix: level.matrix,
        bbox: level.bbox,
        streamInspection,
        path: Object.freeze([level]),
        ...(formInstanceIsolation ? { formInstanceIsolation } : {}),
        ...(instanced ? { formInstance: String(level.doOffset) } : {}),
      }),
    );
  }
  // Forms drawn inside a Form the page shows: each drawing is written on a
  // clone of the chain that reaches it, renamed in its parent when the parent
  // draws it more than once. The parent must name it in resources of its own
  // (a direct dictionary the clone can extend); a Form entered twice on one
  // way is a cycle and is not followed.
  if (pageWritable) {
    // What a Form draws of other Forms, read once per Form however many
    // drawings of it there are (a chart's marker is drawn thousands of times).
    type NestedReading = Readonly<{
      taken: ReadonlySet<string>;
      childForms: ReadonlyMap<string, Readonly<{ objectNumber: number; generation: number }>>;
      childNameCounts: ReadonlyMap<string, number>;
      childInvocations: readonly NativeImageInvocation[];
      childCounts: ReadonlyMap<string, number>;
    }>;
    const nestedReadings = new Map<string, NestedReading | null>();
    const childInspectionsByForm = new Map<string, PdfNativeContentStreamInspection | null>();
    const ownsResources = new Map<string, boolean>();
    const readNested = async (parent: PdfNativeVectorFormHost): Promise<NestedReading | null> => {
      const parentKey = `${parent.form.objectNumber}:${parent.form.generation}`;
      if (nestedReadings.has(parentKey)) return nestedReadings.get(parentKey)!;
      let reading: NestedReading | null = null;
      try {
        const parentObject = await inspectMraiNativeObject(sourceBytes, parent.form);
        const childMappings = parseDirectXObjectDictionaryEntries(
          parentObject.objectBytes,
          "nested-resources",
          "Form",
        );
        const childForms = new Map<string, Readonly<{ objectNumber: number; generation: number }>>();
        const childNameCounts = new Map<string, number>();
        for (const mapping of childMappings) {
          if (mapping.inlineDictionary !== undefined) continue;
          const childDictionary = readMraiNativeObjectDictionarySource(sourceBytes, mapping.target);
          if (childDictionary === null) continue;
          const subtype = parseMraiNativeTopLevelDictionaryNameEntries(latin1Bytes(childDictionary)).filter(
            (entry) => entry.key === "Subtype",
          );
          if (subtype.length !== 1 || subtype[0].directName !== "Form") continue;
          childForms.set(mapping.resourceName, mapping.target);
          const key = `${mapping.target.objectNumber}:${mapping.target.generation}`;
          childNameCounts.set(key, (childNameCounts.get(key) ?? 0) + 1);
        }
        if (childForms.size > 0) {
          const parentStream = parent.streamInspection.streamBytes;
          const childInvocations = scanNativeImageInvocations(parentStream, [
            { stream: Object.freeze({ ...parent.form }), start: 0, end: parentStream.length },
          ]);
          const childCounts = new Map<string, number>();
          for (const invocation of childInvocations) {
            childCounts.set(invocation.resourceName, (childCounts.get(invocation.resourceName) ?? 0) + 1);
          }
          reading = Object.freeze({
            taken: new Set(childMappings.map((entry) => entry.resourceName)),
            childForms,
            childNameCounts,
            childInvocations,
            childCounts,
          });
        }
      } catch (error) {
        if (!(error instanceof PdfNativeTransactionError)) throw error;
        reading = null;
      }
      nestedReadings.set(parentKey, reading);
      return reading;
    };
    for (let cursor = 0; cursor < hosts.length; cursor += 1) {
      const parent = hosts[cursor];
      if (parent.path.length >= MAX_VECTOR_FORM_DEPTH) continue;
      if (resourcesLink) {
        const firstKey = `${parent.path[0].form.objectNumber}:${parent.path[0].form.generation}`;
        let owns = ownsResources.get(firstKey);
        if (owns === undefined) {
          owns = formOwnsResources(sourceBytes, parent.path[0].form);
          ownsResources.set(firstKey, owns);
        }
        if (!owns) continue;
      }
      const reading = await readNested(parent);
      if (!reading) continue;
      const { taken, childForms, childNameCounts, childInvocations, childCounts } = reading;
      const parentChain =
        parent.formInstanceIsolation?.chain ??
        Object.freeze([
          ...resourcesLinkEdges,
          Object.freeze({
            resourceName: parent.resourceName,
            kind: "xobject" as const,
            objectNumber: parent.form.objectNumber,
            generation: parent.form.generation,
            expectedObjectSha256: parent.streamInspection.objectSha256,
          }),
        ]);
      if (parentChain.length >= 8 || !writable(parent.form, true)) continue;
      for (const invocation of childInvocations) {
        const child = childForms.get(invocation.resourceName);
        if (!child) continue;
        const childKey = `${child.objectNumber}:${child.generation}`;
        if (
          childNameCounts.get(childKey) !== 1 ||
          parent.path.some((level) => sameNativeObjectRef(level.form, child)) ||
          !writable(child, true)
        ) {
          continue;
        }
        const childDictionary = readMraiNativeObjectDictionarySource(sourceBytes, child);
        if (childDictionary === null) continue;
        const matrix = formMatrixFromDictionary(childDictionary);
        if (matrix === null) continue;
        let childInspection = childInspectionsByForm.get(childKey);
        if (childInspection === undefined) {
          try {
            childInspection = await inspectMraiNativeContentStream(sourceBytes, child);
          } catch (error) {
            if (!(error instanceof PdfNativeTransactionError)) throw error;
            childInspection = null;
          }
          childInspectionsByForm.set(childKey, childInspection);
        }
        if (!childInspection) continue;
        const instanced = childCounts.get(invocation.resourceName)! > 1;
        let instance: PdfNativeFormInstanceIsolationInstance | undefined;
        if (instanced) {
          const resourceName = freshName(taken, invocation.operatorSpan.start);
          if (!resourceName) continue;
          instance = Object.freeze({
            contentStream: Object.freeze({ ...parent.form }),
            operandStart: invocation.operatorSpan.start,
            resourceName,
          });
        }
        const level: PdfNativeVectorFormHostLevel = Object.freeze({
          resourceName: invocation.resourceName,
          form: Object.freeze({ ...child }),
          doOffset: invocation.operatorSpan.start,
          matrix,
          bbox: formBBoxFromDictionary(childDictionary),
          streamInspection: childInspection,
        });
        const path = Object.freeze([...parent.path, level]);
        hosts.push(
          Object.freeze({
            resourceName: level.resourceName,
            form: level.form,
            doDrawingOffset: parent.doDrawingOffset,
            matrix: level.matrix,
            bbox: level.bbox,
            streamInspection: childInspection,
            path,
            formInstanceIsolation: Object.freeze({
              page: await pageIsolationPin(),
              chain: Object.freeze([
                ...parentChain,
                Object.freeze({
                  resourceName: invocation.resourceName,
                  kind: "xobject" as const,
                  objectNumber: child.objectNumber,
                  generation: child.generation,
                  expectedObjectSha256: childInspection.objectSha256,
                  ...(instance ? { instance } : {}),
                }),
              ]),
            }),
            ...(parent.formInstance !== undefined || instanced
              ? { formInstance: path.map((item) => item.doOffset).join(".") }
              : {}),
          }),
        );
      }
    }
  }
  // A Form reached by more than one host is shown more than once: each of
  // its hosts names its drawing, so their proofs stay apart.
  const hostsPerForm = new Map<string, number>();
  for (const host of hosts) {
    const key = `${host.form.objectNumber}:${host.form.generation}`;
    hostsPerForm.set(key, (hostsPerForm.get(key) ?? 0) + 1);
  }
  return Object.freeze(
    hosts.map((host) =>
      host.formInstance === undefined &&
      hostsPerForm.get(`${host.form.objectNumber}:${host.form.generation}`)! > 1
        ? Object.freeze({
            ...host,
            formInstance: host.path.map((item) => item.doOffset).join("."),
          })
        : host,
    ),
  );
}

/**
 * The XObject names the other pages that resolve names through `owner` draw
 * (their own content streams, scanned the way this page's are), or null when
 * one of those drawings cannot be read -- which counts as drawing everything.
 */
async function resourceNamesDrawnByOtherPages(input: {
  readonly sourceBytes: Uint8Array;
  readonly pages: readonly PdfNativePageContentGraphInspection[];
  readonly pageNumber: number;
  readonly owner: Readonly<{ objectNumber: number; generation: number }>;
}): Promise<ReadonlySet<string> | null> {
  const names = new Set<string>();
  for (const other of input.pages) {
    if (
      other.pageNumber === input.pageNumber ||
      !sameNativeObjectRef(other.resourceAuthority.owner, input.owner)
    ) {
      continue;
    }
    try {
      const inspections = await inspectMraiNativePageContentStreams(
        input.sourceBytes,
        other.pageNumber,
      );
      const drawing = joinNativePageDrawing(inspections);
      const segments = inspections.map((inspection) => ({
        stream: inspection.contentStream,
        start: inspection.drawingOffset,
        end: inspection.drawingOffset + inspection.streamBytes.length,
      }));
      for (const invocation of scanNativeImageInvocations(drawing, segments)) {
        names.add(invocation.resourceName);
      }
    } catch (error) {
      if (error instanceof PdfNativeTransactionError) return null;
      throw error;
    }
  }
  return names;
}

async function nativeImageFormHostInvocations(input: {
  readonly sourceBytes: Uint8Array;
  readonly form: Readonly<{ objectNumber: number; generation: number }>;
  readonly parentOwner: Readonly<{ objectNumber: number; generation: number }>;
  readonly parentKind: PdfNativeResourceDictionaryKind;
  readonly parentOwnerBytes: Uint8Array;
  readonly parentOwnerExclusive: boolean;
  readonly invocationCount: number;
  readonly chain: readonly Readonly<{ objectNumber: number; generation: number }>[];
  readonly orderPrefix: readonly number[];
  /** The page CTM at the form's `Do` (null when unknown). */
  readonly parentMatrix: PdfNativeEmbeddedImageTransformMatrix | null;
}): Promise<Readonly<{
  invocations: readonly NativeHostedImageInvocation[];
  hosts: readonly NativeImageFormHost[];
}> | null> {
  const { sourceBytes, form } = input;
  if (
    input.chain.length >= MAX_NATIVE_IMAGE_FORM_DEPTH ||
    input.chain.some((item) => sameNativeObjectRef(item, form)) ||
    !input.parentOwnerExclusive ||
    input.invocationCount !== 1
  ) {
    return null;
  }
  try {
    if (countMraiNativeIndirectObjectReferences(sourceBytes, form) !== 1) {
      return null;
    }
    const dictionary = readMraiNativeObjectDictionarySource(sourceBytes, form);
    if (dictionary === null) return null;
    const dictionaryBytes = latin1Bytes(dictionary);
    const subtype = parseMraiNativeTopLevelDictionaryNameEntries(
      dictionaryBytes,
    ).filter((entry) => entry.key === "Subtype");
    if (subtype.length !== 1 || subtype[0].directName !== "Form") return null;
    const inspected = await inspectMraiNativeObject(sourceBytes, form);
    if (!inspected.streamBytes) return null;
    const authority = nativeFormResourceAuthority(dictionary);
    if (authority === null) return null;
    let owner = input.parentOwner;
    let kind = input.parentKind;
    let ownerBytes = input.parentOwnerBytes;
    let ownerExclusive: boolean = input.parentOwnerExclusive;
    if (authority.kind === "nested-resources") {
      owner = form;
      kind = "nested-resources";
      ownerBytes = dictionaryBytes;
      ownerExclusive = true;
    } else if (authority.kind === "resource-dictionary") {
      const resources = readMraiNativeObjectDictionarySource(
        sourceBytes,
        authority.ref,
      );
      if (resources === null) return null;
      owner = authority.ref;
      kind = "resource-dictionary";
      ownerBytes = latin1Bytes(resources);
      ownerExclusive =
        countMraiNativeIndirectObjectReferences(sourceBytes, authority.ref) === 1;
    }
    const images = new Map<string, Readonly<{ objectNumber: number; generation: number }>>();
    const forms = new Map<string, Readonly<{ objectNumber: number; generation: number }>>();
    const skipped = new Set<string>();
    for (const mapping of xObjectDictionaryEntries(
      sourceBytes,
      ownerBytes,
      kind,
      "PDF Form",
      true,
    )) {
      const targetDictionary = readMraiNativeObjectDictionarySource(
        sourceBytes,
        mapping.target,
      );
      if (targetDictionary === null) return null;
      const names = parseMraiNativeTopLevelDictionaryNameEntries(
        latin1Bytes(targetDictionary),
      ).filter((entry) => entry.key === "Subtype");
      if (names.length !== 1) return null;
      if (names[0].directName === "Image") images.set(mapping.resourceName, mapping.target);
      else if (names[0].directName === "Form") forms.set(mapping.resourceName, mapping.target);
      else if (names[0].directName === "PS") skipped.add(mapping.resourceName);
      else return null;
    }
    const host: NativeImageFormHost = Object.freeze({
      form: Object.freeze({ ...form }),
      resourceOwner: Object.freeze({ ...owner }),
      resourceKind: kind,
      images,
      chain: Object.freeze([...input.chain, Object.freeze({ ...form })]),
      orderPrefix: input.orderPrefix,
    });
    const scanned = scanNativeImageInvocations(inspected.streamBytes, [
      Object.freeze({ stream: form, start: 0, end: inspected.streamBytes.length }),
    ]);
    // A form draws in the space its /Matrix maps onto the CTM at its `Do`
    // (PDF 32000-1 8.10.1), so its paints' page CTMs are composed from both.
    const formMatrix = formMatrixFromDictionary(dictionary);
    const formSpace =
      input.parentMatrix && formMatrix
        ? concatPdfNativeEmbeddedImageTransformMatrices(formMatrix, input.parentMatrix)
        : null;
    const invocations: NativeHostedImageInvocation[] = [];
    const hosts: NativeImageFormHost[] = [host];
    for (const scannedInvocation of scanned) {
      const invocation: NativeImageInvocation = Object.freeze({
        ...scannedInvocation,
        matrix:
          formSpace && scannedInvocation.matrix
            ? concatPdfNativeEmbeddedImageTransformMatrices(scannedInvocation.matrix, formSpace)
            : null,
      });
      const orderKey = Object.freeze([
        ...input.orderPrefix,
        invocation.operatorSpan.start,
      ]);
      const nestedForm = forms.get(invocation.resourceName);
      if (nestedForm) {
        const nested = await nativeImageFormHostInvocations({
          sourceBytes,
          form: nestedForm,
          parentOwner: owner,
          parentKind: kind,
          parentOwnerBytes: ownerBytes,
          parentOwnerExclusive: ownerExclusive,
          invocationCount: scanned.filter(
            (item) => item.resourceName === invocation.resourceName,
          ).length,
          chain: host.chain,
          orderPrefix: orderKey,
          parentMatrix: invocation.matrix,
        });
        if (nested) {
          invocations.push(...nested.invocations);
          hosts.push(...nested.hosts);
        }
        continue;
      }
      if (skipped.has(invocation.resourceName)) continue;
      // A name the form does not map is a drawing this route cannot read.
      if (!images.has(invocation.resourceName)) return null;
      invocations.push(Object.freeze({ ...invocation, host, orderKey }));
    }
    return Object.freeze({
      invocations: Object.freeze(invocations),
      hosts: Object.freeze(hosts),
    });
  } catch (error) {
    if (error instanceof PdfNativeTransactionError) return null;
    throw error;
  }
}

async function buildNativeImageRemovalPlan(input: {
  readonly sourceBytes: Uint8Array;
  readonly graph: Awaited<ReturnType<typeof resolveNativeImageRemovalGraph>>;
  readonly invocation: NativeHostedImageInvocation;
  readonly image: Readonly<{ objectNumber: number; generation: number }>;
  readonly imageOrder: number;
  readonly drawOperatorBytes: Uint8Array;
  readonly removedOperatorBytes: Uint8Array;
  readonly replacementOperatorBytes: Uint8Array;
  readonly resourceUnmap?: PdfNativePageXObjectUnmapPlan;
}): Promise<PdfNativeEmbeddedImageRemovalPlan> {
  const { graph, invocation, image } = input;
  assertNativeImagePaintStream(graph, invocation);
  const untouched = nativeImagePaintPins(graph, invocation, image);
  const inspected = await inspectMraiNativeContentStream(
    input.sourceBytes,
    invocation.stream,
    untouched.map((item) => ({ ...item, role: "resource" as const })),
  );
  const current = inspected.streamBytes.slice(
    invocation.operatorSpan.start,
    invocation.operatorSpan.end,
  );
  if (
    !bytesEqual(current, input.drawOperatorBytes) &&
    !bytesEqual(current, input.removedOperatorBytes)
  ) {
    fail("precondition-mismatch", "Embedded image paint operator changed.");
  }
  if (
    input.replacementOperatorBytes.length !== current.length ||
    (!bytesEqual(input.replacementOperatorBytes, input.drawOperatorBytes) &&
      !bytesEqual(
        input.replacementOperatorBytes,
        input.removedOperatorBytes,
      )) ||
    bytesEqual(input.replacementOperatorBytes, current)
  ) {
    fail("invalid-request", "Embedded image history bytes are invalid.");
  }
  const evidence = Object.freeze({
    schemaVersion: MRAI_NATIVE_EMBEDDED_IMAGE_REMOVE_SCHEMA_VERSION,
    pageNumber: graph.page.pageNumber,
    imageOrder: input.imageOrder,
    resourceName: invocation.resourceName,
    page: Object.freeze({ ...graph.page.page }),
    resourceOwner: Object.freeze({
      ...nativeImageResourceOwner(graph, invocation.host),
    }),
    image: Object.freeze({ ...image }),
    contentStream: Object.freeze({ ...invocation.stream }),
    operatorSpan: Object.freeze({ ...invocation.operatorSpan }),
    drawOperatorBytes: Object.freeze(Array.from(input.drawOperatorBytes)),
    removedOperatorBytes: Object.freeze(Array.from(input.removedOperatorBytes)),
    ...(input.resourceUnmap
      ? {
          resourceUnmap: Object.freeze({
            resourceName: input.resourceUnmap.resourceName,
            direction: input.resourceUnmap.direction,
            resourceOwner: Object.freeze({
              ...input.resourceUnmap.resourceOwner,
            }),
            expectedOwnerSha256: input.resourceUnmap.expectedOwnerSha256,
            revisedOwnerSha256: input.resourceUnmap.revisedOwnerSha256,
            entryOffset: input.resourceUnmap.entryOffset,
            entryBytes: Object.freeze(
              Array.from(input.resourceUnmap.entryBytes),
            ),
          }),
        }
      : {}),
  });
  return Object.freeze({
    evidence,
    request: Object.freeze({
      identity: inspected.identity,
      expectedObjectSha256: inspected.objectSha256,
      expectedStreamSha256: inspected.streamSha256,
      operatorSpan: Object.freeze({
        ...invocation.operatorSpan,
        expectedSha256: await sha256PdfBytes(current),
      }),
      replacement: Uint8Array.from(input.replacementOperatorBytes),
      untouchedObjects: input.resourceUnmap
        ? inspected.untouchedObjects.filter(
            (item) =>
              !sameNativeObjectRef(item, input.resourceUnmap!.resourceOwner),
          )
        : inspected.untouchedObjects,
      ...(input.resourceUnmap
        ? { pageXObjectUnmap: input.resourceUnmap }
        : {}),
    }),
  });
}

/**
 * Locate one top-level image paint operation by its pdf.js-visible order and
 * build an equal-length native removal. Image resources remain byte-exact;
 * only the selected `/<resource> Do` instance is removed.
 */
export async function planMraiNativeEmbeddedImageRemoval(input: {
  readonly sourceBytes: Uint8Array;
  readonly selection: PdfNativeEmbeddedImageRemovalSelection;
}): Promise<PdfNativeEmbeddedImageRemovalPlan> {
  const pageNumber = exactPositiveInteger(
    input.selection.pageNumber,
    "embedded image page number",
  );
  const imageOrder = exactPositiveInteger(
    input.selection.imageOrder,
    "embedded image order",
    true,
  );
  const graph = await resolveNativeImageRemovalGraph(
    input.sourceBytes,
    pageNumber,
  );
  const invocation = graph.invocations[imageOrder];
  if (!invocation) {
    fail("invalid-request", "Selected embedded image is out of range.");
  }
  assertNativeImagePaintMatrix(invocation, input.selection.matrix);
  const image = nativeImageOf(graph, invocation)!;
  const stream = await inspectMraiNativeContentStream(
    input.sourceBytes,
    invocation.stream,
  );
  const drawOperatorBytes = stream.streamBytes.slice(
    invocation.operatorSpan.start,
    invocation.operatorSpan.end,
  );
  const removedOperatorBytes = new Uint8Array(drawOperatorBytes.length).fill(
    0x20,
  );
  return buildNativeImageRemovalPlan({
    sourceBytes: input.sourceBytes,
    graph,
    invocation,
    image,
    imageOrder,
    drawOperatorBytes,
    removedOperatorBytes,
    replacementOperatorBytes: removedOperatorBytes,
    ...(await lastPaintResourceUnmap(
      input.sourceBytes,
      graph,
      invocation,
      pageNumber,
    )),
  });
}

/**
 * When this removal takes the page's last paint of a picture, drop the
 * picture's name from the page's own resource dictionary as well. Without it
 * the image object stays reachable, so every later export carries its bytes
 * forward; a page that shares its resources with another page keeps them.
 */
async function lastPaintResourceUnmap(
  sourceBytes: Uint8Array,
  graph: Awaited<ReturnType<typeof resolveNativeImageRemovalGraph>>,
  invocation: NativeHostedImageInvocation,
  pageNumber: number,
): Promise<Readonly<{ resourceUnmap?: PdfNativePageXObjectUnmapPlan }>> {
  // The unmap rewrites the page's own resource dictionary; a picture a form
  // host names keeps its entry, and its bytes simply stay in the file.
  if (invocation.host) return Object.freeze({});
  const remaining = graph.invocations.filter(
    (item) =>
      item !== invocation &&
      item.host === null &&
      item.resourceName === invocation.resourceName,
  );
  if (remaining.length > 0) return Object.freeze({});
  try {
    return Object.freeze({
      resourceUnmap: await planMraiNativePageXObjectUnmap({
        sourceBytes,
        pageNumber,
        resourceName: invocation.resourceName,
        direction: "unmap",
      }),
    });
  } catch (error) {
    // Shared or indirect page resources stay exactly as they are; the picture
    // is still removed from the page, its bytes simply remain in the file.
    if (error instanceof PdfNativeTransactionError) return Object.freeze({});
    throw error;
  }
}

/** Rebind one receipt-authenticated remove/restore against current bytes. */
export async function planMraiNativeEmbeddedImageRemovalHistory(input: {
  readonly sourceBytes: Uint8Array;
  readonly expectedSourceSha256: string;
  readonly priorEvidence: PdfNativeEmbeddedImageRemovalEvidence;
  readonly expectedCurrentOperatorBytes: readonly number[];
  readonly desiredOperatorBytes: readonly number[];
}): Promise<PdfNativeEmbeddedImageRemovalPlan> {
  if (
    (await sha256PdfBytes(input.sourceBytes)) !==
    exactSha256(input.expectedSourceSha256, "embedded image history source")
  ) {
    fail("source-mismatch", "Embedded image history source changed.");
  }
  const evidence = input.priorEvidence;
  if (
    evidence.schemaVersion !==
      MRAI_NATIVE_EMBEDDED_IMAGE_REMOVE_SCHEMA_VERSION ||
    !PDF_RESOURCE_NAME.test(evidence.resourceName)
  ) {
    fail("invalid-request", "Embedded image history evidence is invalid.");
  }
  const graph = await resolveNativeImageRemovalGraph(
    input.sourceBytes,
    evidence.pageNumber,
  );
  // The paint may be gone (that is what a restore replays), so its place in
  // the page's order is the count of paints a reader draws before its spot.
  const host =
    graph.formHosts.find((item) => sameNativeObjectRef(item.form, evidence.contentStream)) ??
    null;
  const evidenceOrderKey = nativeImageOrderKeyAt(
    graph,
    evidence.contentStream,
    evidence.operatorSpan.start,
  );
  const precedingInvocationCount =
    evidenceOrderKey === null
      ? -1
      : graph.invocations.filter(
          (invocation) =>
            compareNativeImageOrderKeys(invocation.orderKey, evidenceOrderKey) < 0,
        ).length;
  const hostImages = host ? host.images : graph.images;
  const current = Uint8Array.from(input.expectedCurrentOperatorBytes);
  const desired = Uint8Array.from(input.desiredOperatorBytes);
  const draw = Uint8Array.from(evidence.drawOperatorBytes);
  const removed = Uint8Array.from(evidence.removedOperatorBytes);
  // A record that unmapped the picture's name leaves the page without it
  // while the paint is gone, so the expected mapping follows the direction
  // this replay is about to write.
  const restoring = bytesEqual(desired, draw);
  const mappedAsExpected =
    evidence.resourceUnmap && restoring
      ? !host && !hostImages.has(evidence.resourceName)
      : sameNativeObjectRef(
          hostImages.get(evidence.resourceName) ?? {
            objectNumber: 0,
            generation: 0,
          },
          evidence.image,
        );
  if (
    evidenceOrderKey === null ||
    precedingInvocationCount !== evidence.imageOrder ||
    !sameNativeObjectRef(graph.page.page, evidence.page) ||
    !sameNativeObjectRef(
      nativeImageResourceOwner(graph, host),
      evidence.resourceOwner,
    ) ||
    !mappedAsExpected
  ) {
    fail("precondition-mismatch", "Embedded image source graph changed.");
  }
  if (
    current.length === 0 ||
    current.length !== desired.length ||
    current.length !== draw.length ||
    (!bytesEqual(current, draw) && !bytesEqual(current, removed)) ||
    (!bytesEqual(desired, draw) && !bytesEqual(desired, removed)) ||
    bytesEqual(current, desired)
  ) {
    fail("invalid-request", "Embedded image history direction is invalid.");
  }
  const invocation: NativeHostedImageInvocation = Object.freeze({
    resourceName: evidence.resourceName,
    stream: evidence.contentStream,
    operatorSpan: evidence.operatorSpan,
    // A replay is bound by its recorded bytes, never by a reader's measure.
    matrix: null,
    host,
    orderKey: evidenceOrderKey!,
  });
  // Restoring the paint puts the name back; removing it again drops it.
  const resourceUnmap = evidence.resourceUnmap
    ? await planMraiNativePageXObjectUnmap({
        sourceBytes: input.sourceBytes,
        pageNumber: evidence.pageNumber,
        resourceName: evidence.resourceName,
        direction: restoring ? "remap" : "unmap",
        ...(restoring
          ? {
              entryOffset: evidence.resourceUnmap.entryOffset,
              entryBytes: Uint8Array.from(evidence.resourceUnmap.entryBytes),
            }
          : {}),
      })
    : undefined;
  return buildNativeImageRemovalPlan({
    sourceBytes: input.sourceBytes,
    graph,
    invocation,
    image: evidence.image,
    imageOrder: evidence.imageOrder,
    drawOperatorBytes: draw,
    removedOperatorBytes: removed,
    replacementOperatorBytes: desired,
    ...(resourceUnmap ? { resourceUnmap } : {}),
  });
}

function normalizedEmbeddedImageReplacementData(
  value: PdfNativeEmbeddedImageReplacementData,
): Readonly<{
  evidence: PdfNativeEmbeddedImageReplacementData;
  encodedBytes: Uint8Array;
  softMaskBytes: Uint8Array | null;
}> {
  const width = exactReplacementImageDimension(value.width, "width");
  const height = exactReplacementImageDimension(value.height, "height");
  if (
    width * height > MAX_MRAI_NATIVE_EMBEDDED_IMAGE_PIXELS ||
    value.bitsPerComponent !== 8 ||
    (value.colorSpace !== "DeviceGray" && value.colorSpace !== "DeviceRGB") ||
    (value.filter !== "DCTDecode" && value.filter !== "FlateDecode") ||
    (value.mimeType === "image/jpeg") !== (value.filter === "DCTDecode") ||
    (value.mimeType === "image/png") !== (value.filter === "FlateDecode")
  ) {
    fail("invalid-request", "Replacement image authority is invalid.");
  }
  const colors = value.colorSpace === "DeviceGray" ? 1 : 3;
  if (
    (value.filter === "DCTDecode" && value.decodeParms !== null) ||
    (value.filter === "FlateDecode" &&
      (!value.decodeParms ||
        value.decodeParms.predictor !== 15 ||
        value.decodeParms.colors !== colors ||
        value.decodeParms.bitsPerComponent !== 8 ||
        value.decodeParms.columns !== width))
  ) {
    fail("invalid-request", "Replacement image decode authority is invalid.");
  }
  const encodedBytes = base64ToBytes(value.encodedDataBase64);
  const encodedSha256 = exactSha256(
    value.encodedSha256,
    "replacement image payload",
  );
  let softMaskBytes: Uint8Array | null = null;
  if (value.softMask !== undefined) {
    if (
      !value.softMask ||
      value.mimeType !== "image/png" ||
      width * height * 4 > MAX_DECODED_STREAM_BYTES
    ) {
      fail("invalid-request", "Replacement image alpha authority is invalid.");
    }
    softMaskBytes = base64ToBytes(value.softMask.encodedDataBase64);
    const decodedMask = decodeFlateStream(softMaskBytes);
    if (
      decodedMask.length !== (width + 1) * height ||
      encodedBytes.length + softMaskBytes.length >
        MAX_MRAI_NATIVE_EMBEDDED_IMAGE_BYTES
    ) {
      fail("invalid-request", "Replacement image alpha size is invalid.");
    }
    for (let row = 0; row < height; row += 1) {
      if (decodedMask[row * (width + 1)] > 4)
        fail(
          "invalid-request",
          "Replacement image alpha row filter is invalid.",
        );
    }
    exactSha256(value.softMask.encodedSha256, "replacement alpha payload");
  }
  return Object.freeze({
    evidence: Object.freeze({
      mimeType: value.mimeType,
      width,
      height,
      colorSpace: value.colorSpace,
      bitsPerComponent: 8,
      filter: value.filter,
      decodeParms: value.decodeParms
        ? Object.freeze({ ...value.decodeParms })
        : null,
      encodedDataBase64: value.encodedDataBase64,
      encodedSha256,
      ...(value.softMask
        ? { softMask: Object.freeze({ ...value.softMask }) }
        : {}),
    }),
    encodedBytes,
    softMaskBytes,
  });
}

function serializeEmbeddedImageObject(
  reference: Readonly<{ objectNumber: number; generation: number }>,
  data: PdfNativeEmbeddedImageReplacementData,
  encodedBytes: Uint8Array,
  softMaskRef?: Readonly<{ objectNumber: number; generation: number }>,
): Uint8Array {
  const decodeParms = data.decodeParms
    ? ` /DecodeParms << /Predictor 15 /Colors ${data.decodeParms.colors} /BitsPerComponent 8 /Columns ${data.decodeParms.columns} >>`
    : "";
  const softMask = softMaskRef
    ? ` /SMask ${softMaskRef.objectNumber} ${softMaskRef.generation} R`
    : "";
  return concatBytes([
    ascii(
      `${reference.objectNumber} ${reference.generation} obj\n<< /Type /XObject /Subtype /Image /Width ${data.width} /Height ${data.height} /ColorSpace /${data.colorSpace} /BitsPerComponent 8 /Filter /${data.filter}${decodeParms}${softMask} /Length ${encodedBytes.length} >>\nstream\n`,
    ),
    encodedBytes,
    ascii("\nendstream\nendobj"),
  ]);
}














async function validateReplaceableSourceImage(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
  reference: Readonly<{ objectNumber: number; generation: number }>,
  allowSoftMask = true,
): Promise<ParsedIndirectObject> {
  const image = parseIndirectObject(
    sourceBytes,
    parsed,
    reference.objectNumber,
    reference.generation,
  );
  exactNamedObjectType(image.bytes, "Subtype", "Image", "PDF image");
  if (
    image.storage !== "direct" ||
    image.dictionaryStart === null ||
    image.dictionaryEnd === null ||
    image.encodedStreamBytes === null ||
    image.streamBytes === null
  ) {
    // In an object stream: kept as it is.
    return image;
  }
  const source = asLatin1(image.bytes);
  exactNamedObjectType(image.bytes, "Type", "XObject", "PDF image", true);
  const entries = topLevelDictionaryEntries(source);
  if (
    entries.some((entry) =>
      new Set(["Mask", "Alternates", "OPI", "Matte"]).has(entry.key),
    )
  ) {
    // A masked or alternate picture: kept as it is.
    return image;
  }
  const width = oneTopLevelInteger(source, "Width", "PDF image Width");
  const height = oneTopLevelInteger(source, "Height", "PDF image Height");
  const bits = oneTopLevelInteger(
    source,
    "BitsPerComponent",
    "PDF image BitsPerComponent",
  );
  const colorSpaceEntry = oneTopLevelDictionaryEntry(source, "ColorSpace", "PDF image ColorSpace");
  if (!colorSpaceEntry || source[colorSpaceEntry.valueStart] !== "/") {
    // A color space array or reference (ICCBased, Indexed, ...): kept as it is.
    return image;
  }
  const colorSpace = oneTopLevelName(
    source,
    "ColorSpace",
    "PDF image ColorSpace",
  );
  if (
    width < 1 ||
    height < 1 ||
    width * height > MAX_MRAI_NATIVE_EMBEDDED_IMAGE_PIXELS
  ) {
    fail("unsupported-document", "PDF image color authority is unsupported.");
  }
  if (bits !== 8 || (colorSpace !== "DeviceGray" && colorSpace !== "DeviceRGB")) {
    // Another depth or device space (CMYK, 1-bit, 16-bit): kept as it is.
    return image;
  }
  const sourceMask = oneTopLevelDictionaryEntry(
    source,
    "SMask",
    "PDF image soft mask",
    true,
  );
  if (sourceMask) {
    if (!allowSoftMask)
      fail("unsupported-document", "Nested PDF image masks are unsupported.");
    const maskRef = parseTopLevelReference(
      source,
      sourceMask.valueStart,
      sourceMask.valueEnd,
      "PDF image soft mask",
    );
    if (
      maskRef.end !== sourceMask.valueEnd ||
      sameNativeObjectRef(maskRef.reference, reference)
    ) {
      fail("unsupported-document", "PDF image soft mask reference is invalid.");
    }
    const mask = await validateReplaceableSourceImage(
      sourceBytes,
      parsed,
      maskRef.reference,
      false,
    );
    const maskSource = asLatin1(mask.bytes);
    if (
      oneTopLevelName(maskSource, "ColorSpace", "PDF mask color space") !==
        "DeviceGray" ||
      oneTopLevelInteger(maskSource, "Width", "PDF mask width") !== width ||
      oneTopLevelInteger(maskSource, "Height", "PDF mask height") !== height
    ) {
      fail(
        "unsupported-document",
        "PDF image and soft mask dimensions disagree.",
      );
    }
  }
  const filter = image.streamEncoding;
  if (
    !filter ||
    !["identity", "flate", "flate-predictor15", "dct", "ascii85-dct"].includes(
      filter,
    )
  ) {
    // A filter this engine does not decode (JPX, CCITT, JBIG2): kept as it is.
    return image;
  }
  const decodeParms = entries.filter((entry) => entry.key === "DecodeParms");
  const colors = colorSpace === "DeviceGray" ? 1 : 3;
  if (image.streamEncoding === "flate-predictor15") {
    const parameters = source.slice(
      decodeParms[0].valueStart,
      decodeParms[0].valueEnd,
    );
    const predictor = imagePredictorParameters(parameters);
    const rowLength = width * colors + 1;
    if (
      !predictor ||
      predictor.columns !== width ||
      predictor.colors !== colors ||
      predictor.bits !== bits ||
      image.streamBytes.length !== rowLength * height
    ) {
      fail(
        "unsupported-document",
        "Source image predictor scanlines are inconsistent.",
      );
    }
    for (let row = 0; row < height; row += 1) {
      if (image.streamBytes[row * rowLength] > 4) {
        fail("unsupported-document", "Source image row filter is invalid.");
      }
    }
    return image;
  }
  // parseStreamEncoding already limits other DecodeParms to neutral Predictor 1.
  if (decodeParms.length > 0 && image.streamEncoding !== "flate") {
    fail(
      "unsupported-document",
      "Source image decode parameters are unsupported.",
    );
  }
  if (
    filter === "identity" &&
    image.streamBytes.length !== width * height * colors
  ) {
    fail("unsupported-document", "Raw source image length is unsupported.");
  }
  if (
    filter === "flate" &&
    image.streamBytes.length !== width * height * colors
  ) {
    fail("unsupported-document", "Flate source image length is unsupported.");
  }
  if (filter === "dct" || filter === "ascii85-dct") {
    const jpeg = await replacementImageDataFromJpeg(image.streamBytes);
    if (
      jpeg.width !== width ||
      jpeg.height !== height ||
      jpeg.colorSpace !== colorSpace
    ) {
      fail(
        "unsupported-document",
        "JPEG source image authority is inconsistent.",
      );
    }
  }
  return image;
}

function directPageXObjectDictionarySpan(
  objectBytes: Uint8Array,
): Readonly<{ start: number; end: number }> {
  const source = asLatin1(objectBytes);
  const resources = oneTopLevelDictionaryEntry(
    source,
    "Resources",
    "PDF page Resources",
  );
  if (
    !resources ||
    source.slice(resources.valueStart, resources.valueStart + 2) !== "<<" ||
    dictionaryEnd(source, resources.valueStart) !== resources.valueEnd
  ) {
    fail(
      "unsupported-document",
      "Inherited or indirect page Resources are preserve-only.",
    );
  }
  const resourceSource = source.slice(resources.valueStart, resources.valueEnd);
  const xObject = oneTopLevelDictionaryEntry(
    resourceSource,
    "XObject",
    "PDF page XObject resources",
  );
  if (
    !xObject ||
    resourceSource.slice(xObject.valueStart, xObject.valueStart + 2) !== "<<" ||
    dictionaryEnd(resourceSource, xObject.valueStart) !== xObject.valueEnd
  ) {
    fail(
      "unsupported-document",
      "Indirect page XObject resources are preserve-only.",
    );
  }
  return Object.freeze({
    start: resources.valueStart + xObject.valueStart,
    end: resources.valueStart + xObject.valueEnd,
  });
}

/** The /XObject dictionary of a page's own resources object (the page says
 * `/Resources N G R`), written in that object as a direct dictionary. */
function resourceObjectXObjectDictionarySpan(
  objectBytes: Uint8Array,
): Readonly<{ start: number; end: number }> {
  const source = asLatin1(objectBytes);
  const xObject = oneTopLevelDictionaryEntry(
    source,
    "XObject",
    "PDF page XObject resources",
  );
  if (
    !xObject ||
    source.slice(xObject.valueStart, xObject.valueStart + 2) !== "<<" ||
    dictionaryEnd(source, xObject.valueStart) !== xObject.valueEnd
  ) {
    fail(
      "unsupported-document",
      "Indirect page XObject resources are preserve-only.",
    );
  }
  return Object.freeze({ start: xObject.valueStart, end: xObject.valueEnd });
}

function rewritePageXObjectMapping(input: {
  readonly pageObjectBytes: Uint8Array;
  readonly resourceName: string;
  readonly reference: Readonly<{ objectNumber: number; generation: number }>;
  readonly operation: "add" | "remove";
  /** The bytes are the page's own resources object, not the page. */
  readonly resourcesObject?: boolean;
}): Uint8Array {
  if (!PDF_RESOURCE_NAME.test(input.resourceName)) {
    fail("invalid-request", "Replacement image resource name is invalid.");
  }
  const source = asLatin1(input.pageObjectBytes);
  const span = input.resourcesObject
    ? resourceObjectXObjectDictionarySpan(input.pageObjectBytes)
    : directPageXObjectDictionarySpan(input.pageObjectBytes);
  const dictionary = source.slice(span.start, span.end);
  const entries = topLevelDictionaryEntries(dictionary);
  const matches = entries.filter((entry) => entry.key === input.resourceName);
  if (input.operation === "add") {
    if (matches.length !== 0) {
      fail(
        "precondition-mismatch",
        "Replacement image resource already exists.",
      );
    }
    const insertion = `\n/${input.resourceName} ${input.reference.objectNumber} ${input.reference.generation} R`;
    return concatBytes([
      input.pageObjectBytes.slice(0, span.end - 2),
      ascii(insertion),
      input.pageObjectBytes.slice(span.end - 2),
    ]);
  }
  if (matches.length !== 1) {
    fail("precondition-mismatch", "Replacement image resource is missing.");
  }
  const match = matches[0];
  const parsed = parseTopLevelReference(
    dictionary,
    match.valueStart,
    match.valueEnd,
    "replacement image resource",
  );
  if (
    parsed.end !== match.valueEnd ||
    !sameNativeObjectRef(parsed.reference, input.reference) ||
    match !== entries.at(-1)
  ) {
    fail("precondition-mismatch", "Replacement image resource changed.");
  }
  const absoluteStart = span.start + match.keyStart;
  const removalStart =
    absoluteStart > span.start + 2 && source[absoluteStart - 1] === "\n"
      ? absoluteStart - 1
      : absoluteStart;
  return concatBytes([
    input.pageObjectBytes.slice(0, removalStart),
    input.pageObjectBytes.slice(span.start + match.valueEnd),
  ]);
}

function nextEmbeddedImageResourceName(names: ReadonlySet<string>): string {
  for (let index = 0; index < 10_000; index += 1) {
    const candidate = `MRAIImg${index}`;
    if (!names.has(candidate)) return candidate;
  }
  fail("unsupported-document", "PDF image resource namespace is exhausted.");
}

function reviseContentStreamObject(input: {
  readonly sourceBytes: Uint8Array;
  readonly object: ParsedIndirectObject;
  readonly start: number;
  readonly end: number;
  readonly replacement: Uint8Array;
}): Readonly<{
  bytes: Uint8Array;
  encodedStream: Uint8Array;
  decodedStream: Uint8Array;
}> {
  const object = input.object;
  assertTargetStream(object);
  const decodedStream = concatBytes([
    object.streamBytes.slice(0, input.start),
    input.replacement,
    object.streamBytes.slice(input.end),
  ]);
  const encodedStream =
    object.streamEncoding === "flate"
      ? encodeDeterministicFlateStream(decodedStream)
      : object.streamEncoding === "ascii85-flate"
        ? encodeAscii85Stream(encodeDeterministicFlateStream(decodedStream))
        : object.streamEncoding === "ascii-hex"
          ? encodeAsciiHexStream(decodedStream)
          : object.streamEncoding === "lzw"
            ? encodeLzwStream(decodedStream)
            : decodedStream;
  return Object.freeze({
    bytes: rewriteStreamObjectBytes(input.sourceBytes, object, encodedStream),
    encodedStream,
    decodedStream,
  });
}

/** Whether the page's own /Resources entry names `owner` (not a dictionary
 * it inherits from a parent Pages node). */
function pageNamesResourcesObject(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
  pageRef: Readonly<{ objectNumber: number; generation: number }>,
  owner: Readonly<{ objectNumber: number; generation: number }>,
): boolean {
  const pageObject = parseIndirectObject(
    sourceBytes,
    parsed,
    pageRef.objectNumber,
    pageRef.generation,
  );
  const pageSource = asLatin1(pageObject.bytes);
  const resources = oneTopLevelDictionaryEntry(
    pageSource,
    "Resources",
    "PDF page Resources",
    true,
  );
  if (!resources || pageSource.slice(resources.valueStart, resources.valueStart + 2) === "<<") {
    return false;
  }
  const named = parseTopLevelReference(
    pageSource,
    resources.valueStart,
    resources.valueEnd,
    "PDF page Resources",
  );
  return named.end === resources.valueEnd && sameNativeObjectRef(named.reference, owner);
}

async function buildNativeEmbeddedImageReplacement(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly imageOrder: number;
  readonly replacement: PdfNativeEmbeddedImageReplacementData;
  readonly priorEvidence?: PdfNativeEmbeddedImageReplacementEvidence;
  readonly expectedCurrent: "original" | "replacement";
  readonly desired: "original" | "replacement";
  readonly expectedMatrix?: PdfNativeEmbeddedImageTransformMatrix | null;
}): Promise<PdfNativePreparedEmbeddedImageReplacement> {
  if (input.expectedCurrent === input.desired) {
    fail("invalid-request", "Embedded image replacement no-op is forbidden.");
  }
  const parsed = parsePdf(input.sourceBytes, false);
  const graph = await resolveNativeImageRemovalGraph(
    input.sourceBytes,
    input.pageNumber,
  );
  const invocation = graph.invocations[input.imageOrder];
  if (!invocation) {
    fail("invalid-request", "Selected embedded image is out of range.");
  }
  assertNativeImagePaintMatrix(invocation, input.expectedMatrix);
  // A replacement rewrites the page's own XObject entry; a picture a form
  // host paints is named in the form's resources instead.
  if (invocation.host) {
    fail(
      "unsupported-document",
      "Replacing a picture a Form XObject paints is preserve-only.",
    );
  }
  // The new name goes where the page keeps its names: in its own /Resources
  // dictionary, or in a resources object the page names as its own (most
  // producers write it so, and so do MRAI's own font and opacity installs;
  // a picture on such a page could not be replaced, user audit PDF 0048).
  // Names a parent Pages node lends, or one object several pages read, are
  // never written.
  const authority = graph.page.resourceAuthority;
  const ownerIsPage =
    authority.dictionaryKind === "nested-resources" &&
    sameNativeObjectRef(graph.page.page, authority.owner);
  const ownResourcesObject =
    !ownerIsPage &&
    authority.dictionaryKind === "resource-dictionary" &&
    !sameNativeObjectRef(graph.page.page, authority.owner) &&
    pageNamesResourcesObject(input.sourceBytes, parsed, graph.page.page, authority.owner) &&
    countMraiNativeIndirectObjectReferences(input.sourceBytes, authority.owner) === 1;
  if (!ownerIsPage && !ownResourcesObject) {
    fail(
      "unsupported-document",
      "Inherited or shared page Resources are preserve-only.",
    );
  }
  const page = parseIndirectObject(
    input.sourceBytes,
    parsed,
    authority.owner.objectNumber,
    authority.owner.generation,
  );
  if (page.storage !== "direct" || page.streamBytes !== null) {
    fail("unsupported-document", "Compressed PDF pages are preserve-only.");
  }
  const contentKey = `${invocation.stream.objectNumber}:${invocation.stream.generation}`;
  if (graph.contentStreamReferenceCounts.get(contentKey) !== 1) {
    fail("unsupported-document", "Embedded image content stream is shared.");
  }
  const normalized = normalizedEmbeddedImageReplacementData(input.replacement);
  if (
    (await sha256PdfBytes(normalized.encodedBytes)) !==
    normalized.evidence.encodedSha256
  ) {
    fail("invalid-request", "Replacement image payload hash changed.");
  }
  if (
    normalized.softMaskBytes &&
    (await sha256PdfBytes(normalized.softMaskBytes)) !==
      normalized.evidence.softMask!.encodedSha256
  ) {
    fail("invalid-request", "Replacement image alpha hash changed.");
  }
  const latest = latestEntries(parsed);
  const highestObjectNumber = Math.max(0, ...latest.keys());
  const originalResourceName =
    input.priorEvidence?.originalResourceName ?? invocation.resourceName;
  const originalImage =
    input.priorEvidence?.originalImage ??
    graph.images.get(originalResourceName) ??
    fail("precondition-mismatch", "Original image resource is missing.");
  await validateReplaceableSourceImage(
    input.sourceBytes,
    parsed,
    originalImage,
  );
  const replacementResourceName =
    input.priorEvidence?.replacementResourceName ??
    nextEmbeddedImageResourceName(new Set(graph.images.keys()));
  const replacementRef =
    input.priorEvidence?.replacementImage ??
    Object.freeze({ objectNumber: highestObjectNumber + 1, generation: 0 });
  assertObjectReference(replacementRef.objectNumber, replacementRef.generation);
  if (
    input.priorEvidence &&
    Boolean(input.priorEvidence.replacementSoftMask) !==
      Boolean(normalized.softMaskBytes)
  ) {
    fail("precondition-mismatch", "Replacement image alpha authority changed.");
  }
  const maskRef = normalized.softMaskBytes
    ? (input.priorEvidence?.replacementSoftMask ?? {
        objectNumber: replacementRef.objectNumber + 1,
        generation: 0,
      })
    : undefined;
  if (maskRef) assertObjectReference(maskRef.objectNumber, maskRef.generation);
  if (
    maskRef &&
    (maskRef.objectNumber !== replacementRef.objectNumber + 1 ||
      maskRef.generation !== 0)
  ) {
    fail("invalid-request", "Replacement soft mask allocation is invalid.");
  }
  const maskObject =
    maskRef && normalized.softMaskBytes
      ? serializeEmbeddedImageObject(
          maskRef,
          {
            ...normalized.evidence,
            colorSpace: "DeviceGray",
            filter: "FlateDecode",
            decodeParms: {
              predictor: 15,
              colors: 1,
              bitsPerComponent: 8,
              columns: normalized.evidence.width,
            },
          },
          normalized.softMaskBytes,
        )
      : null;
  const maskObjectSha256 = maskObject ? await sha256PdfBytes(maskObject) : null;
  const replacementObject = serializeEmbeddedImageObject(
    replacementRef,
    normalized.evidence,
    normalized.encodedBytes,
    maskRef,
  );
  const replacementObjectSha256 = await sha256PdfBytes(replacementObject);
  if (
    input.priorEvidence &&
    input.priorEvidence.replacementImage.expectedObjectSha256 !==
      replacementObjectSha256
  ) {
    fail(
      "precondition-mismatch",
      "Replacement image object authority changed.",
    );
  }
  const expectedResourceName =
    input.expectedCurrent === "original"
      ? originalResourceName
      : replacementResourceName;
  if (invocation.resourceName !== expectedResourceName) {
    fail("precondition-mismatch", "Embedded image visible instance changed.");
  }
  const originalMapping = graph.images.get(originalResourceName);
  if (
    !originalMapping ||
    !sameNativeObjectRef(originalMapping, originalImage)
  ) {
    fail("precondition-mismatch", "Original image resource identity changed.");
  }
  const replacementMapping = graph.images.get(replacementResourceName) ?? null;
  if (
    (input.expectedCurrent === "replacement" &&
      (!replacementMapping ||
        !sameNativeObjectRef(replacementMapping, replacementRef))) ||
    (input.expectedCurrent === "original" && replacementMapping !== null)
  ) {
    fail("precondition-mismatch", "Replacement image mapping state changed.");
  }
  if (input.priorEvidence) {
    const existingReplacement = parseIndirectObject(
      input.sourceBytes,
      parsed,
      replacementRef.objectNumber,
      replacementRef.generation,
    );
    if (
      existingReplacement.storage !== "direct" ||
      (await sha256PdfBytes(existingReplacement.bytes)) !==
        replacementObjectSha256
    ) {
      fail("precondition-mismatch", "Replacement image object changed.");
    }
    if (maskRef && maskObjectSha256) {
      const existingMask = parseIndirectObject(
        input.sourceBytes,
        parsed,
        maskRef.objectNumber,
        maskRef.generation,
      );
      if (
        input.priorEvidence.replacementSoftMask?.expectedObjectSha256 !==
          maskObjectSha256 ||
        existingMask.storage !== "direct" ||
        (await sha256PdfBytes(existingMask.bytes)) !== maskObjectSha256
      ) {
        fail("precondition-mismatch", "Replacement soft mask object changed.");
      }
    }
  } else if (latest.get(replacementRef.objectNumber)?.inUse) {
    fail(
      "precondition-mismatch",
      "Replacement image object allocation changed.",
    );
  }
  if (
    !input.priorEvidence &&
    maskRef &&
    latest.get(maskRef.objectNumber)?.inUse
  ) {
    fail("precondition-mismatch", "Replacement soft mask allocation changed.");
  }
  const content = parseIndirectObject(
    input.sourceBytes,
    parsed,
    invocation.stream.objectNumber,
    invocation.stream.generation,
  );
  assertTargetStream(content);
  const currentOperator = content.streamBytes.slice(
    invocation.operatorSpan.start,
    invocation.operatorSpan.end,
  );
  const expectedOperator =
    input.expectedCurrent === "original"
      ? input.priorEvidence
        ? Uint8Array.from(input.priorEvidence.originalOperatorBytes)
        : currentOperator
      : Uint8Array.from(input.priorEvidence!.replacementOperatorBytes);
  if (!bytesEqual(currentOperator, expectedOperator)) {
    fail("precondition-mismatch", "Embedded image paint operator changed.");
  }
  const desiredOperator =
    input.desired === "original"
      ? Uint8Array.from(input.priorEvidence!.originalOperatorBytes)
      : ascii(`/${replacementResourceName} Do`);
  const revisedContent = reviseContentStreamObject({
    sourceBytes: input.sourceBytes,
    object: content,
    start: invocation.operatorSpan.start,
    end: invocation.operatorSpan.end,
    replacement: desiredOperator,
  });
  const revisedPage = rewritePageXObjectMapping({
    pageObjectBytes: page.bytes,
    resourceName: replacementResourceName,
    reference: replacementRef,
    operation: input.desired === "replacement" ? "add" : "remove",
    resourcesObject: ownResourcesObject,
  });
  const evidence: PdfNativeEmbeddedImageReplacementEvidence =
    input.priorEvidence
      ? Object.freeze({
          ...input.priorEvidence,
          replacement: normalized.evidence,
        })
      : Object.freeze({
          schemaVersion: MRAI_NATIVE_EMBEDDED_IMAGE_REPLACE_SCHEMA_VERSION,
          pageNumber: graph.page.pageNumber,
          imageOrder: input.imageOrder,
          originalResourceName,
          replacementResourceName,
          page: Object.freeze({ ...graph.page.page }),
          resourceOwner: Object.freeze({
            ...graph.page.resourceAuthority.owner,
          }),
          originalImage: Object.freeze({ ...originalImage }),
          replacementImage: Object.freeze({
            ...replacementRef,
            expectedObjectSha256: replacementObjectSha256,
          }),
          ...(maskRef && maskObjectSha256
            ? {
                replacementSoftMask: Object.freeze({
                  ...maskRef,
                  expectedObjectSha256: maskObjectSha256,
                }),
              }
            : {}),
          contentStream: Object.freeze({ ...invocation.stream }),
          originalOperatorBytes: Object.freeze(Array.from(currentOperator)),
          replacementOperatorBytes: Object.freeze(Array.from(desiredOperator)),
          replacement: normalized.evidence,
        });
  if (
    !sameNativeObjectRef(evidence.page, graph.page.page) ||
    !sameNativeObjectRef(
      evidence.resourceOwner,
      graph.page.resourceAuthority.owner,
    ) ||
    !sameNativeObjectRef(evidence.contentStream, invocation.stream) ||
    evidence.pageNumber !== input.pageNumber ||
    evidence.imageOrder !== input.imageOrder
  ) {
    fail(
      "precondition-mismatch",
      "Embedded image replacement evidence changed.",
    );
  }
  const revisions: IncrementalObjectRevision[] = [
    Object.freeze({
      objectNumber: content.objectNumber,
      generation: content.generation,
      bytes: revisedContent.bytes,
    }),
    Object.freeze({
      objectNumber: page.objectNumber,
      generation: page.generation,
      bytes: revisedPage,
    }),
    ...(!input.priorEvidence
      ? [
          Object.freeze({
            objectNumber: replacementRef.objectNumber,
            generation: replacementRef.generation,
            bytes: replacementObject,
          }),
          ...(maskRef && maskObject
            ? [
                Object.freeze({
                  objectNumber: maskRef.objectNumber,
                  generation: maskRef.generation,
                  bytes: maskObject,
                }),
              ]
            : []),
        ]
      : []),
  ].sort(
    (left, right) =>
      left.objectNumber - right.objectNumber ||
      left.generation - right.generation,
  );
  const incremental = makeMultiObjectIncrementalRevision(
    input.sourceBytes,
    parsed.current,
    revisions,
    { min: 2, max: maskRef ? 4 : 3, label: "Embedded image replacement" },
  );
  const revisedObjects = await Promise.all(
    revisions.map(async (revision) =>
      Object.freeze({
        objectNumber: revision.objectNumber,
        generation: revision.generation,
        offset: incremental.offsets.get(
          `${revision.objectNumber}:${revision.generation}`,
        )!,
        sha256: await sha256PdfBytes(revision.bytes),
      }),
    ),
  );
  // Over a source that claims PDF/A the catalog's packet is written last in
  
  const pdfaIdentification = await appendPdfaIdentificationRemoval(input.sourceBytes, parsed, incremental);
  // A picture with an alpha channel in a file declaring PDF 1.3: the
  // catalog states 1.4, last in the same revision (user audit PDF 0020).
  const catalogVersion = await appendCatalogVersionRaise(input.sourceBytes, parsed, pdfaIdentification ?? incremental);
  const candidateBytes = catalogVersion?.candidateBytes ?? pdfaIdentification?.candidateBytes ?? incremental.candidateBytes;
  for (const beside of [pdfaIdentification?.removal, catalogVersion?.raise]) {
    if (!beside) continue;
    const { objectNumber, generation, offset, sha256 } = beside;
    revisedObjects.push(Object.freeze({ objectNumber, generation, offset, sha256 }));
    revisedObjects.sort((left, right) => left.objectNumber - right.objectNumber || left.generation - right.generation);
  }
  return Object.freeze({
    evidence,
    sourceLength: input.sourceBytes.length,
    sourceXrefOffset: parsed.current.offset,
    incrementalXrefOffset:
      catalogVersion?.incrementalXrefOffset ?? pdfaIdentification?.incrementalXrefOffset ?? incremental.incrementalXrefOffset,
    candidateSha256: await sha256PdfBytes(candidateBytes),
    candidateBytes,
    revisedObjects: Object.freeze(revisedObjects),
    ...(pdfaIdentification ? { pdfaIdentificationRemoval: pdfaIdentification.removal } : {}),
    ...(catalogVersion ? { catalogVersionRaise: catalogVersion.raise } : {}),
  });
}

const EMBEDDED_IMAGE_TRANSFORM_SEARCH_WINDOW = 512;

// The block's opening, searched backwards from the draw: `q`, the rectangle,
// its clip and the inner level up to the draw.
const EMBEDDED_IMAGE_FRAME_OPENING = new RegExp(
  `(?:^|[ \\t\\r\\n\\f\\0])q${EMBEDDED_IMAGE_FRAME_WS}(${EMBEDDED_IMAGE_TRANSFORM_NUMBER}${EMBEDDED_IMAGE_FRAME_WS}` +
    `${EMBEDDED_IMAGE_TRANSFORM_NUMBER}${EMBEDDED_IMAGE_FRAME_WS}${EMBEDDED_IMAGE_TRANSFORM_NUMBER}${EMBEDDED_IMAGE_FRAME_WS}` +
    `${EMBEDDED_IMAGE_TRANSFORM_NUMBER}${EMBEDDED_IMAGE_FRAME_MIDDLE})$`,
);
const EMBEDDED_IMAGE_FRAME_CLOSING = new RegExp(
  `^${EMBEDDED_IMAGE_FRAME_WS}Q${EMBEDDED_IMAGE_FRAME_WS}Q(?:$|[ \\t\\r\\n\\f\\0])`,
);
const EMBEDDED_IMAGE_FRAME_TOLERANCE = 0.05;

export interface PdfNativeEmbeddedImageFrameClip {
  /** The clip rectangle as written, `x y w h` in the block's own space. */
  readonly rect: readonly [number, number, number, number];
  readonly rectText: string;
  /** From the white space after the rectangle to the draw, byte for byte. */
  readonly middle: string;
  /** The inner level's `cm`: the picture's placement inside the block. */
  readonly innerMatrix: PdfNativeEmbeddedImageTransformMatrix;
  /** The paths the block clips the picture by after its rectangle (a shape
   * the picture is cut to), as written: they move with the picture. */
  readonly clipPaths?: string;
  /** The draw the block holds: a bare `/X Do` or one of our wrappers. */
  readonly drawText: string;
}


/** The box the picture covers in the block's space: the unit square (or the
 * wrapper's clip of it) through the wrapper, then the inner `cm`. */
function embeddedImageFrameBox(
  innerMatrix: PdfNativeEmbeddedImageTransformMatrix,
  wrapper: Readonly<{
    matrix: PdfNativeEmbeddedImageTransformMatrix;
    clip: PdfNativeEmbeddedImageTransformClip | null;
  }> | null,
): readonly [number, number, number, number] {
  const apply = (m: PdfNativeEmbeddedImageTransformMatrix, x: number, y: number) =>
    [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]] as const;
  const clip = wrapper?.clip ?? { x: 0, y: 0, width: 1, height: 1 };
  const xs: number[] = [];
  const ys: number[] = [];
  for (const [u, v] of [
    [clip.x, clip.y],
    [clip.x + clip.width, clip.y],
    [clip.x, clip.y + clip.height],
    [clip.x + clip.width, clip.y + clip.height],
  ] as const) {
    const [px, py] = wrapper ? apply(wrapper.matrix, u, v) : [u, v];
    const [x, y] = apply(innerMatrix, px, py);
    xs.push(x);
    ys.push(y);
  }
  const x = Math.min(...xs);
  const y = Math.min(...ys);
  return [x, y, Math.max(...xs) - x, Math.max(...ys) - y];
}

function embeddedImageFrameRectMatches(
  rect: readonly [number, number, number, number],
  box: readonly [number, number, number, number],
): boolean {
  return rect.every(
    (value, index) => Math.abs(value - box[index]) <= EMBEDDED_IMAGE_FRAME_TOLERANCE,
  );
}

type EmbeddedImageFrameWrapper = Readonly<{
  matrix: PdfNativeEmbeddedImageTransformMatrix;
  clip: PdfNativeEmbeddedImageTransformClip | null;
}>;

function embeddedImageFrameInvert(
  m: readonly number[],
): readonly [number, number, number, number, number, number] {
  const det = m[0] * m[3] - m[1] * m[2];
  return [
    m[3] / det,
    -m[1] / det,
    -m[2] / det,
    m[0] / det,
    (m[2] * m[5] - m[3] * m[4]) / det,
    (m[1] * m[4] - m[0] * m[5]) / det,
  ];
}

/**
 * Where a frame clip goes when its draw changes from `before` to `after`.
 * A clip that is the picture's box -- as it stands, or where the producer
 * placed it before an earlier move left the clip behind -- becomes the new
 * box; any other clip (a crop of the picture) moves with the picture.
 */
export function pdfNativeEmbeddedImageFrameCarriedRect(
  frame: Pick<PdfNativeEmbeddedImageFrameClip, "rect" | "innerMatrix">,
  before: EmbeddedImageFrameWrapper | null,
  after: EmbeddedImageFrameWrapper,
): readonly [number, number, number, number] {
  if (
    embeddedImageFrameRectMatches(frame.rect, embeddedImageFrameBox(frame.innerMatrix, before)) ||
    embeddedImageFrameRectMatches(frame.rect, embeddedImageFrameBox(frame.innerMatrix, null))
  ) {
    return embeddedImageFrameBox(frame.innerMatrix, after);
  }
  const apply = (m: readonly number[], x: number, y: number) =>
    [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]] as const;
  const innerInverse = embeddedImageFrameInvert(frame.innerMatrix);
  const beforeInverse = before ? embeddedImageFrameInvert(before.matrix) : null;
  const [rx, ry, rw, rh] = frame.rect;
  const xs: number[] = [];
  const ys: number[] = [];
  for (const [px, py] of [
    [rx, ry],
    [rx + rw, ry],
    [rx, ry + rh],
    [rx + rw, ry + rh],
  ] as const) {
    let [x, y] = apply(innerInverse, px, py);
    if (beforeInverse) [x, y] = apply(beforeInverse, x, y);
    [x, y] = apply(after.matrix, x, y);
    [x, y] = apply(frame.innerMatrix, x, y);
    xs.push(x);
    ys.push(y);
  }
  const x = Math.min(...xs);
  const y = Math.min(...ys);
  return [x, y, Math.max(...xs) - x, Math.max(...ys) - y];
}

/**
 * The block's clip paths (a shape the picture is cut to) moved as the
 * picture moves: every point through the same change the rectangle takes
 * (pdfNativeEmbeddedImageFrameCarriedRect), a rectangle path as the four
 * corners it turns into. Null when a path does not read as points and path
 * operators.
 */
export function pdfNativeEmbeddedImageFrameCarriedPaths(
  frame: Pick<PdfNativeEmbeddedImageFrameClip, "clipPaths" | "innerMatrix">,
  before: EmbeddedImageFrameWrapper | null,
  after: EmbeddedImageFrameWrapper,
): string | null {
  const paths = frame.clipPaths;
  if (!paths) return "";
  const apply = (m: readonly number[], x: number, y: number) =>
    [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]] as const;
  const innerInverse = embeddedImageFrameInvert(frame.innerMatrix);
  const beforeInverse = before ? embeddedImageFrameInvert(before.matrix) : null;
  const carry = (px: number, py: number) => {
    let [x, y] = apply(innerInverse, px, py);
    if (beforeInverse) [x, y] = apply(beforeInverse, x, y);
    [x, y] = apply(after.matrix, x, y);
    return apply(frame.innerMatrix, x, y);
  };
  const number = formatPdfNativeEmbeddedImageTransformNumber;
  const point = (x: number, y: number) => { const [cx, cy] = carry(x, y); return `${number(cx)} ${number(cy)}`; };
  const tokens = paths.trim().split(/[ \t\r\n\f\0]+/u);
  const out: string[] = [];
  let operands: number[] = [];
  const counts: Record<string, number> = { m: 2, l: 2, c: 6, v: 4, y: 4, h: 0, re: 4 };
  for (const token of tokens) {
    if (/^[-+]?(?:\d+(?:\.\d*)?|\.\d+)$/u.test(token)) {
      operands.push(Number(token));
      continue;
    }
    if (token === "W" || token === "W*" || token === "n") {
      if (operands.length) return null;
      out.push(token);
      continue;
    }
    const count = counts[token];
    if (count === undefined || operands.length !== count || operands.some((value) => !Number.isFinite(value))) return null;
    if (token === "re") {
      const [x, y, w, h] = operands;
      out.push(`${point(x, y)} m ${point(x + w, y)} l ${point(x + w, y + h)} l ${point(x, y + h)} l h`);
    } else {
      const pairs: string[] = [];
      for (let index = 0; index < operands.length; index += 2) pairs.push(point(operands[index], operands[index + 1]));
      out.push([...pairs, token].join(" "));
    }
    operands = [];
  }
  if (operands.length) return null;
  return "\n" + out.join(" ").replace(/ n /gu, " n\n");
}

/** Widen a bare-draw or wrapper span to the Skia frame clip block that holds
 * it, when there is one and its rectangle is the picture's box. */
function extendEmbeddedImageSpanToFrameClip(
  streamBytes: Uint8Array,
  span: Readonly<{ start: number; end: number }>,
  resourceName: string,
): Readonly<{ start: number; end: number }> {
  const closing = decodeLatin1(streamBytes.slice(span.end, span.end + 64));
  if (!EMBEDDED_IMAGE_FRAME_CLOSING.test(closing)) return span;
  const windowStart = Math.max(0, span.start - EMBEDDED_IMAGE_TRANSFORM_SEARCH_WINDOW);
  const opening = EMBEDDED_IMAGE_FRAME_OPENING.exec(
    decodeLatin1(streamBytes.slice(windowStart, span.start)),
  );
  if (!opening) return span;
  const start = span.start - opening[1].length;
  const frame = parsePdfNativeEmbeddedImageFrameClip(
    decodeLatin1(streamBytes.slice(start, span.end)),
  );
  if (!frame) return span;
  const wrapper = parsePdfNativeEmbeddedImageTransformOperator(frame.drawText);
  if (
    wrapper
      ? wrapper.resourceName !== resourceName
      : !isPdfNativeEmbeddedImageBareDraw(frame.drawText, resourceName)
  ) {
    return span;
  }
  // The block holds this one draw, so its clip belongs to this picture.
  return Object.freeze({ start, end: span.end });
}


/** Six-decimal PDF number without trailing zeros, never "-0". */
export function formatPdfNativeEmbeddedImageTransformNumber(
  value: number,
): string {
  if (!Number.isFinite(value)) {
    fail("invalid-request", "Embedded image transform value is not finite.");
  }
  const rounded = Math.round(value * 1_000_000) / 1_000_000;
  const text = (Object.is(rounded, -0) ? 0 : rounded)
    .toFixed(6)
    .replace(/\.?0+$/, "");
  return text === "" || text === "-" ? "0" : text;
}



export function buildPdfNativeEmbeddedImageTransformOperator(input: {
  readonly resourceName: string;
  readonly matrix: PdfNativeEmbeddedImageTransformMatrix;
  readonly clip: PdfNativeEmbeddedImageTransformClip | null;
  /** Page ExtGState carrying this instance's constant alpha, when set. */
  readonly graphicsStateName?: string | null;
}): string {
  if (!PDF_RESOURCE_NAME.test(input.resourceName)) {
    fail("invalid-request", "Embedded image resource name is invalid.");
  }
  if (
    input.graphicsStateName != null &&
    !PDF_RESOURCE_NAME.test(input.graphicsStateName)
  ) {
    fail("invalid-request", "Embedded image graphics state name is invalid.");
  }
  const fmt = formatPdfNativeEmbeddedImageTransformNumber;
  const state = input.graphicsStateName ? `/${input.graphicsStateName} gs ` : "";
  const clip = input.clip
    ? `${fmt(input.clip.x)} ${fmt(input.clip.y)} ${fmt(input.clip.width)} ${fmt(input.clip.height)} re W n `
    : "";
  return `q ${state}${input.matrix.map(fmt).join(" ")} cm ${clip}/${input.resourceName} Do Q`;
}


/**
 * A bare draw of one picture: its name and `Do`, separated by any run of PDF
 * white space (ISO 32000-1 7.2.2), as producers write it on one line or two.
 */
export function isPdfNativeEmbeddedImageBareDraw(
  text: string,
  resourceName: string,
): boolean {
  if (!text.startsWith(`/${resourceName}`) || !text.endsWith("Do")) return false;
  const gap = text.slice(resourceName.length + 1, text.length - 2);
  return gap.length > 0 && /^[\0\t\n\f\r ]+$/.test(gap);
}

/**
 * The bytes a transform owns: the bare `/<name> Do` the scanner found, or the
 * whole `q … cm … Do Q` wrapper a previous transform wrote around it. A
 * wrapper must be a complete operator string of this writer's own grammar so
 * a later transform replaces it instead of nesting.
 */
function locateEmbeddedImageTransformSpan(
  streamBytes: Uint8Array,
  invocation: NativeImageInvocation,
): Readonly<{ start: number; end: number }> {
  const bare = invocation.operatorSpan;
  const suffixEnd = bare.end + 2;
  if (
    suffixEnd <= streamBytes.length &&
    decodeLatin1(streamBytes.slice(bare.end, suffixEnd)) === " Q"
  ) {
    const windowStart = Math.max(0, bare.start - EMBEDDED_IMAGE_TRANSFORM_SEARCH_WINDOW);
    const window = decodeLatin1(streamBytes.slice(windowStart, suffixEnd));
    const bareOffset = bare.start - windowStart;
    for (let index = bareOffset - 1; index >= 0; index -= 1) {
      if (window[index] !== "q" || window[index + 1] !== " ") continue;
      if (index > 0 && !/[\s]/.test(window[index - 1])) continue;
      const parsed = parsePdfNativeEmbeddedImageTransformOperator(
        window.slice(index),
      );
      if (parsed && parsed.resourceName === invocation.resourceName) {
        return Object.freeze({ start: windowStart + index, end: suffixEnd });
      }
    }
  }
  return Object.freeze({ start: bare.start, end: bare.end });
}

async function resolveEmbeddedImageTransformTarget(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly imageOrder: number;
  /** History written before frame clips were carried spans the draw only;
   * its Undo/Redo keeps that span. */
  readonly withoutFrameClip?: boolean;
}) {
  const graph = await resolveNativeImageRemovalGraph(
    input.sourceBytes,
    input.pageNumber,
  );
  const invocation = graph.invocations[input.imageOrder];
  if (!invocation) {
    fail("invalid-request", "Selected embedded image is out of range.");
  }
  const image = nativeImageOf(graph, invocation);
  if (!image) {
    fail("precondition-mismatch", "Embedded image resource is missing.");
  }
  assertNativeImagePaintStream(graph, invocation);
  const untouched = nativeImagePaintPins(graph, invocation, image);
  const inspected = await inspectMraiNativeContentStream(
    input.sourceBytes,
    invocation.stream,
    untouched.map((item) => ({ ...item, role: "resource" as const })),
  );
  const drawSpan = locateEmbeddedImageTransformSpan(inspected.streamBytes, invocation);
  const span = input.withoutFrameClip
    ? drawSpan
    : extendEmbeddedImageSpanToFrameClip(
        inspected.streamBytes,
        drawSpan,
        invocation.resourceName,
      );
  const current = inspected.streamBytes.slice(span.start, span.end);
  return { graph, invocation, image, inspected, span, current };
}

export interface PdfNativeEmbeddedImagePaintBlock {
  readonly imageOrder: number;
  readonly resourceName: string;
  readonly image: PdfNativeEmbeddedImageObjectRef;
  readonly contentStream: PdfNativeEmbeddedImageObjectRef;
  /** Which of the page's content streams draws this picture; -1 when a form
   * host draws it (`formStreamIndex` then names which). */
  readonly streamIndex: number;
  /** Which of `formStreams` draws this picture, for a picture a Form XObject
   * drawn once paints; its spans are then offsets into that form's stream. */
  readonly formStreamIndex?: number;
  /** Byte span of the paint: the bare `Do` or this writer's whole wrapper. */
  readonly span: Readonly<{ start: number; end: number }>;
  /** The same span, as an offset into the whole page drawing. */
  readonly drawingSpan: Readonly<{ start: number; end: number }>;
  readonly wrapped: boolean;
  /** `/ImageMask true`: the picture paints in the ambient fill colour. */
  readonly stencil: boolean;
}

export interface PdfNativeEmbeddedImagePageStream {
  readonly index: number;
  readonly contentStream: PdfNativeEmbeddedImageObjectRef;
  readonly streamBytes: Uint8Array;
  /** Where this stream sits inside the page drawing, as [start, end). */
  readonly drawingSpan: Readonly<{ start: number; end: number }>;
  readonly untouchedObjects: PdfNativeContentStreamInspection["untouchedObjects"];
  readonly identity: PdfNativeContentStreamInspection["identity"];
  readonly objectSha256: string;
  readonly streamSha256: string;
  /** For a Form XObject host's stream: where that form resolves its names. */
  readonly resourceOwner?: PdfNativeEmbeddedImageObjectRef;
}

export interface PdfNativeEmbeddedImagePagePaints {
  readonly pageNumber: number;
  readonly page: PdfNativeEmbeddedImageObjectRef;
  readonly resourceOwner: PdfNativeEmbeddedImageObjectRef;
  /** Every content stream the page draws from, in /Contents order. */
  readonly streams: readonly PdfNativeEmbeddedImagePageStream[];
  /** The page drawing: its streams joined the way a reader joins them. */
  readonly drawingBytes: Uint8Array;
  /** The stream of every Form XObject host that paints a picture listed
   * here, each its own drawing (`drawingSpan` covers the whole stream). */
  readonly formStreams: readonly PdfNativeEmbeddedImagePageStream[];
  readonly paints: readonly PdfNativeEmbeddedImagePaintBlock[];
}

/**
 * Where each picture paint the writers address on one page is drawn: the CTM
 * in force at the paint, listed by `imageOrder` (null where the walk cannot
 * know it). A reader that measured a picture with pdf.js finds its paint here
 * before a write (`resolvePdfNativeEmbeddedImageOrder`), because the reader's
 * count of paints is not always the writers'.
 */
export async function listMraiNativeEmbeddedImagePaintMatrices(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
}): Promise<readonly (PdfNativeEmbeddedImageTransformMatrix | null)[]> {
  const pageNumber = exactPositiveInteger(
    input.pageNumber,
    "embedded image page number",
  );
  const graph = await resolveNativeImageRemovalGraph(input.sourceBytes, pageNumber);
  return Object.freeze(graph.invocations.map((invocation) => invocation.matrix));
}











export async function inspectMraiNativeEmbeddedImagePagePaints(
  sourceBytes: Uint8Array,
  pageNumber: number,
): Promise<PdfNativeEmbeddedImagePagePaints> {
  const graph = await resolveNativeImageRemovalGraph(sourceBytes, pageNumber);
  // Pin the page, its resource owner and every image the page paints: a
  // reorder moves draw operators, so none of them may change underneath it.
  // The forms that host pictures, where they resolve names and the pictures
  // they paint are pinned the same way.
  const pinned = [
    graph.page.page,
    graph.page.resourceAuthority.owner,
    ...graph.images.values(),
    ...graph.formHosts.flatMap((host) => [
      host.form,
      host.resourceOwner,
      ...host.images.values(),
    ]),
  ].filter(
    (item, index, items) =>
      items.findIndex((candidate) => sameNativeObjectRef(candidate, item)) === index,
  );
  const inspections = await inspectMraiNativePageContentStreams(
    sourceBytes,
    pageNumber,
    {
      extraUntouched: pinned.map((item) => ({ ...item, role: "resource" as const })),
    },
  );
  for (const inspection of inspections) {
    const key = `${inspection.contentStream.objectNumber}:${inspection.contentStream.generation}`;
    if (graph.contentStreamReferenceCounts.get(key) !== 1) {
      fail(
        "unsupported-document",
        "Embedded image content stream is shared across page Contents.",
      );
    }
  }
  const streams = Object.freeze(
    inspections.map((inspection) =>
      Object.freeze({
        index: inspection.streamIndex,
        contentStream: inspection.contentStream,
        streamBytes: inspection.streamBytes,
        drawingSpan: Object.freeze({
          start: inspection.drawingOffset,
          end: inspection.drawingOffset + inspection.streamBytes.length,
        }),
        untouchedObjects: inspection.untouchedObjects,
        identity: inspection.identity,
        objectSha256: inspection.objectSha256,
        streamSha256: inspection.streamSha256,
      }),
    ),
  );
  const drawingBytes = joinNativePageDrawing(inspections);
  // A form host's stream is its own drawing, rewritten on its own; it pins
  // the page's streams and everything the page pins but itself.
  const paintingHosts = graph.formHosts.filter((host) =>
    graph.invocations.some((invocation) => invocation.host === host),
  );
  const formStreams = Object.freeze(
    await Promise.all(
      paintingHosts.map(async (host, index) => {
        const inspection = await inspectMraiNativeContentStream(
          sourceBytes,
          host.form,
          [...pinned, ...graph.page.contentStreams]
            .filter(
              (item, position, items) =>
                !sameNativeObjectRef(item, host.form) &&
                items.findIndex((candidate) => sameNativeObjectRef(candidate, item)) ===
                  position,
            )
            .map((item) => ({ ...item, role: "resource" as const })),
        );
        return Object.freeze({
          index,
          contentStream: Object.freeze({ ...host.form }),
          resourceOwner: Object.freeze({ ...host.resourceOwner }),
          streamBytes: inspection.streamBytes,
          drawingSpan: Object.freeze({ start: 0, end: inspection.streamBytes.length }),
          untouchedObjects: inspection.untouchedObjects,
          identity: inspection.identity,
          objectSha256: inspection.objectSha256,
          streamSha256: inspection.streamSha256,
        });
      }),
    ),
  );
  const stencils = new Map<string, boolean>();
  const stencilKey = (image: Readonly<{ objectNumber: number; generation: number }>) =>
    `${image.objectNumber}:${image.generation}`;
  for (const image of [
    ...graph.images.values(),
    ...graph.formHosts.flatMap((host) => [...host.images.values()]),
  ]) {
    // Whether a paint is a stencil is stated in the picture's dictionary.
    // Reading it must not ask this engine to decode the picture: listing what
    // a page paints has to work for JPX, CCITT, JBIG2 and predictor chains it
    // will never decode, exactly as the removal graph already reads them.
    const dictionary = readMraiNativeObjectDictionarySource(sourceBytes, image);
    if (dictionary === null) {
      fail("unsupported-document", "PDF XObject is not readable.");
    }
    stencils.set(
      stencilKey(image),
      parseMraiNativeTopLevelDictionaryBoolean(
        latin1Bytes(dictionary),
        "ImageMask",
      ) === true,
    );
  }
  const paints = graph.invocations.map((invocation, imageOrder) => {
    const image = nativeImageOf(graph, invocation);
    const stream = (invocation.host ? formStreams : streams).find((candidate) =>
      sameNativeObjectRef(candidate.contentStream, invocation.stream),
    );
    if (!image || !stream) {
      fail("unsupported-document", "Embedded image paint is not addressable.");
    }
    const span = locateEmbeddedImageTransformSpan(
      stream.streamBytes,
      invocation,
    );
    return Object.freeze({
      imageOrder,
      resourceName: invocation.resourceName,
      image: Object.freeze({ ...image }),
      contentStream: stream.contentStream,
      streamIndex: invocation.host ? -1 : stream.index,
      ...(invocation.host ? { formStreamIndex: stream.index } : {}),
      span: Object.freeze({ ...span }),
      drawingSpan: Object.freeze({
        start: stream.drawingSpan.start + span.start,
        end: stream.drawingSpan.start + span.end,
      }),
      wrapped: span.start !== invocation.operatorSpan.start,
      stencil: stencils.get(stencilKey(image)) === true,
    });
  });
  return Object.freeze({
    pageNumber: graph.page.pageNumber,
    page: Object.freeze({ ...graph.page.page }),
    resourceOwner: Object.freeze({ ...graph.page.resourceAuthority.owner }),
    streams,
    drawingBytes,
    formStreams,
    paints: Object.freeze(paints),
  });
}

async function buildNativeEmbeddedImageTransformPlan(input: {
  readonly target: Awaited<ReturnType<typeof resolveEmbeddedImageTransformTarget>>;
  readonly imageOrder: number;
  readonly evidence: PdfNativeEmbeddedImageTransformEvidence;
  readonly desiredOperatorBytes: Uint8Array;
  readonly pageExtGStateMutation?: PdfNativePageExtGStateMutationPlan;
}): Promise<PdfNativeEmbeddedImageTransformPlan> {
  const { target } = input;
  const original = Uint8Array.from(input.evidence.originalOperatorBytes);
  const transformed = Uint8Array.from(input.evidence.transformedOperatorBytes);
  if (
    !bytesEqual(target.current, original) &&
    !bytesEqual(target.current, transformed)
  ) {
    fail("precondition-mismatch", "Embedded image paint operator changed.");
  }
  if (
    (!bytesEqual(input.desiredOperatorBytes, original) &&
      !bytesEqual(input.desiredOperatorBytes, transformed)) ||
    bytesEqual(input.desiredOperatorBytes, target.current)
  ) {
    fail("invalid-request", "Embedded image transform history bytes are invalid.");
  }
  if (
    !sameNativeObjectRef(input.evidence.page, target.graph.page.page) ||
    !sameNativeObjectRef(
      input.evidence.resourceOwner,
      nativeImageResourceOwner(target.graph, target.invocation.host),
    ) ||
    !sameNativeObjectRef(input.evidence.image, target.image) ||
    !sameNativeObjectRef(input.evidence.contentStream, target.invocation.stream) ||
    input.evidence.resourceName !== target.invocation.resourceName ||
    input.evidence.pageNumber !== target.graph.page.pageNumber ||
    input.evidence.imageOrder !== input.imageOrder
  ) {
    fail("precondition-mismatch", "Embedded image transform evidence changed.");
  }
  const mutation = input.pageExtGStateMutation;
  // Objects this revision rewrites cannot also be claimed untouched.
  const mutableResourceIdentities = new Set(
    mutation
      ? [
          `${mutation.page.objectNumber}:${mutation.page.generation}`,
          `${mutation.resourceOwner.objectNumber}:${mutation.resourceOwner.generation}`,
          ...(mutation.category.kind === "indirect"
            ? [
                `${mutation.category.target.objectNumber}:${mutation.category.target.generation}`,
              ]
            : []),
        ]
      : [],
  );
  return Object.freeze({
    evidence: input.evidence,
    request: Object.freeze({
      identity: target.inspected.identity,
      expectedObjectSha256: target.inspected.objectSha256,
      expectedStreamSha256: target.inspected.streamSha256,
      operatorSpan: Object.freeze({
        start: target.span.start,
        end: target.span.end,
        expectedSha256: await sha256PdfBytes(target.current),
      }),
      replacement: Uint8Array.from(input.desiredOperatorBytes),
      untouchedObjects: mutation
        ? Object.freeze(
            target.inspected.untouchedObjects.filter(
              (item) =>
                !mutableResourceIdentities.has(
                  `${item.objectNumber}:${item.generation}`,
                ),
            ),
          )
        : target.inspected.untouchedObjects,
      ...(mutation ? { pageExtGStateMutation: mutation } : {}),
    }),
  });
}

function exactEmbeddedImageOpacity(value: number): number {
  if (
    typeof value !== "number" ||
    !Number.isFinite(value) ||
    value < 0 ||
    value > 1
  ) {
    fail("invalid-request", "Embedded image opacity is invalid.");
  }
  return value;
}

/**
 * Deterministic name for the page ExtGState that carries one image alpha, so
 * repeated edits reuse a single resource instead of appending new ones.
 */
async function pdfNativeEmbeddedImageOpacityResourceName(
  opacity: number,
): Promise<string> {
  const digest = await sha256PdfBytes(
    ascii(
      `mrai.image-opacity.v1|${formatNativePdfNumber(opacity)}`,
    ),
  );
  return `MRAI_IO_${digest.slice(0, 24)}`;
}


export function invertPdfNativeEmbeddedImageTransformMatrix(
  matrix: PdfNativeEmbeddedImageTransformMatrix,
): PdfNativeEmbeddedImageTransformMatrix {
  const [a, b, c, d, e, f] = matrix;
  const det = a * d - b * c;
  if (!Number.isFinite(det) || Math.abs(det) < 1e-9) {
    fail("invalid-request", "Embedded image transform matrix is degenerate.");
  }
  const ia = d / det;
  const ib = -b / det;
  const ic = -c / det;
  const id = a / det;
  return Object.freeze([
    ia,
    ib,
    ic,
    id,
    -(e * ia + f * ic),
    -(e * ib + f * id),
  ]) as PdfNativeEmbeddedImageTransformMatrix;
}

/** On-page frame in user space (bottom-left origin) the visible image fills. */
export interface PdfNativeEmbeddedImageTransformFrame {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export interface PdfNativeEmbeddedImageTransformIntent {
  /** The CTM the reader walked at the paint (user space). */
  readonly ctm: PdfNativeEmbeddedImageTransformMatrix;
  /** Where the visible part of the image must land. */
  readonly frame: PdfNativeEmbeddedImageTransformFrame;
  /** Clockwise quarter turns on screen accumulated by this edit. */
  readonly quarterTurns: 0 | 1 | 2 | 3;
  /** Mirror left/right as shown on screen, after the turns. */
  readonly flipX: boolean;
  /** Mirror top/bottom as shown on screen, after the turns. */
  readonly flipY: boolean;
  /**
   * Retained region of the currently VISIBLE image (fractions, top-left
   * origin), composed onto an existing clip; `null`/absent keeps the existing
   * clip.
   */
  readonly crop?: PdfNativeEmbeddedImageTransformFrame | null;
  /**
   * Constant alpha for this paint instance (0..1). Absent keeps whatever the
   * current wrapper carries; a number installs (or reuses) one page ExtGState
   * resource and names it inside the wrapper.
   */
  readonly opacity?: number | null;
  /**
   * A user-space affine applied after the paint as `ctm` walked it (`cm`
   * order): the reader frame's move, resize along the picture's own axes,
   * free-angle turn or mirror in one matrix. It replaces the frame fit, so
   * the turns must then be 0 and the flips false; the crop and the opacity
   * apply as above. Absent: the frame, turns and flips place the picture.
   */
  readonly pageTransform?: PdfNativeEmbeddedImageTransformMatrix | null;
  /**
   * Word's Crop > Fill, Fit or Reset Crop of the picture's own crop
   * (pdfNativeEmbeddedImagePlacement): the frame, turns, flips, crop and page
   * transform are not given beside it.
   */
  readonly cropMode?: PdfNativeEmbeddedImageCropMode | null;
}





/**
 * Rewrite one image paint instance in place so the visible part of the
 * image lands on `intent.frame` with the requested turns, flips and crop
 * (or is carried by `intent.pageTransform`, at any angle).
 * `intent.ctm` is the CTM in force at the draw as the reader walked it; the
 * wrapper matrix written before the draw is `placement × ctm⁻¹` (an earlier
 * wrapper on the same instance is folded out and replaced, never nested,
 * and its clip is kept unless a new crop composes onto it).
 */
export async function planMraiNativeEmbeddedImageTransform(input: {
  readonly sourceBytes: Uint8Array;
  readonly selection: PdfNativeEmbeddedImageTransformSelection;
  readonly intent: PdfNativeEmbeddedImageTransformIntent;
}): Promise<PdfNativeEmbeddedImageTransformPlan> {
  const pageNumber = exactPositiveInteger(
    input.selection.pageNumber,
    "embedded image page number",
  );
  const imageOrder = exactPositiveInteger(
    input.selection.imageOrder,
    "embedded image order",
    true,
  );
  const intent = input.intent;
  const ctm = exactEmbeddedImageTransformMatrix([...intent.ctm]);
  if (![0, 1, 2, 3].includes(intent.quarterTurns)) {
    fail("invalid-request", "Embedded image transform turns are invalid.");
  }
  const target = await resolveEmbeddedImageTransformTarget({
    sourceBytes: input.sourceBytes,
    pageNumber,
    imageOrder,
  });
  // The wrapper is computed from `ctm`; on another paint it would throw that
  // picture somewhere nobody asked for.
  assertNativeImagePaintMatrix(target.invocation, ctm);
  const currentText = decodeLatin1(target.current);
  const frame = parsePdfNativeEmbeddedImageFrameClip(currentText);
  const drawText = frame ? frame.drawText : currentText;
  const wrapped = parsePdfNativeEmbeddedImageTransformOperator(drawText);
  if (
    !isPdfNativeEmbeddedImageBareDraw(drawText, target.invocation.resourceName) &&
    (!wrapped || wrapped.resourceName !== target.invocation.resourceName)
  ) {
    fail("unsupported-document", "Embedded image draw operator is not transformable.");
  }
  // The reader's CTM already includes an earlier wrapper; fold it out so the
  // new wrapper composes onto the stream's own CTM.
  const outerCtm = wrapped
    ? concatPdfNativeEmbeddedImageTransformMatrices(
        invertPdfNativeEmbeddedImageTransformMatrix(wrapped.matrix),
        ctm,
      )
    : ctm;
  const { placement, clip } = pdfNativeEmbeddedImagePlacement({
    ctm,
    existingClip: wrapped?.clip ?? null,
    frame: intent.frame,
    quarterTurns: intent.quarterTurns,
    flipX: Boolean(intent.flipX),
    flipY: Boolean(intent.flipY),
    crop: intent.crop ?? null,
    pageTransform: intent.pageTransform ?? null,
    cropMode: intent.cropMode ?? null,
  });
  const matrix = exactEmbeddedImageTransformMatrix([
    ...concatPdfNativeEmbeddedImageTransformMatrices(
      placement,
      invertPdfNativeEmbeddedImageTransformMatrix(outerCtm),
    ),
  ]);
  // Alpha travels in a page ExtGState resource, exactly like a vector's
  // opacity: the wrapper names it, so the paint instance carries its own
  // constant alpha without touching the image object or its pixels.
  const opacity =
    intent.opacity === undefined || intent.opacity === null
      ? null
      : exactEmbeddedImageOpacity(intent.opacity);
  // The alpha resource is installed in the page's own ExtGState dictionary,
  // which a form that declares its own Resources does not read.
  if (opacity !== null && target.invocation.host) {
    fail(
      "unsupported-document",
      "Opacity for a picture a Form XObject paints is preserve-only.",
    );
  }
  const graphicsStateName =
    opacity === null
      ? (wrapped?.graphicsStateName ?? null)
      : await pdfNativeEmbeddedImageOpacityResourceName(opacity);
  const pageExtGStateMutation =
    opacity === null
      ? null
      : await planMraiNativePageExtGStateMutation({
          sourceBytes: input.sourceBytes,
          pageNumber,
          resourceName: graphicsStateName!,
          fillOpacity: opacity,
          strokeOpacity: null,
          blendMode: null,
        });
  const transformedDraw = buildPdfNativeEmbeddedImageTransformOperator({
    resourceName: target.invocation.resourceName,
    matrix,
    clip,
    graphicsStateName,
  });
  const carriedPaths = frame ? pdfNativeEmbeddedImageFrameCarriedPaths(frame, wrapped, { matrix, clip }) : "";
  if (carriedPaths === null) {
    fail("unsupported-document", "Embedded image clip path is not transformable.");
  }
  const transformedText = frame
    ? pdfNativeEmbeddedImageFrameCarriedRect(frame, wrapped, { matrix, clip })
        .map(formatPdfNativeEmbeddedImageTransformNumber)
        .join(" ") +
      (frame.clipPaths ? frame.middle.replace(frame.clipPaths, carriedPaths) : frame.middle) +
      transformedDraw
    : transformedDraw;
  if (transformedText.length > EMBEDDED_IMAGE_TRANSFORM_MAX_OPERATOR_BYTES) {
    fail("invalid-request", "Embedded image transform operator is too long.");
  }
  const transformed = ascii(transformedText);
  if (bytesEqual(transformed, target.current)) {
    fail("invalid-request", "Embedded image transform is a no-op.");
  }
  const evidence: PdfNativeEmbeddedImageTransformEvidence = Object.freeze({
    schemaVersion: MRAI_NATIVE_EMBEDDED_IMAGE_TRANSFORM_SCHEMA_VERSION,
    pageNumber: target.graph.page.pageNumber,
    imageOrder,
    resourceName: target.invocation.resourceName,
    page: Object.freeze({ ...target.graph.page.page }),
    resourceOwner: Object.freeze({
      ...nativeImageResourceOwner(target.graph, target.invocation.host),
    }),
    image: Object.freeze({ ...target.image }),
    contentStream: Object.freeze({ ...target.invocation.stream }),
    originalOperatorBytes: Object.freeze(Array.from(target.current)),
    transformedOperatorBytes: Object.freeze(Array.from(transformed)),
    matrix,
    clip,
    ...(graphicsStateName ? { graphicsStateName } : {}),
    ...(opacity === null ? {} : { opacity }),
  });
  return buildNativeEmbeddedImageTransformPlan({
    target,
    imageOrder,
    evidence,
    desiredOperatorBytes: transformed,
    ...(pageExtGStateMutation ? { pageExtGStateMutation } : {}),
  });
}


/** Rebind one receipt-authenticated transform/restore against current bytes. */
export async function planMraiNativeEmbeddedImageTransformHistory(input: {
  readonly sourceBytes: Uint8Array;
  readonly expectedSourceSha256: string;
  readonly priorEvidence: PdfNativeEmbeddedImageTransformEvidence;
  readonly expectedCurrentOperatorBytes: readonly number[];
  readonly desiredOperatorBytes: readonly number[];
}): Promise<PdfNativeEmbeddedImageTransformPlan> {
  if (
    (await sha256PdfBytes(input.sourceBytes)) !==
    exactSha256(input.expectedSourceSha256, "embedded image history source")
  ) {
    fail("source-mismatch", "Embedded image history source changed.");
  }
  const evidence = readPdfNativeEmbeddedImageTransformEvidence(
    input.priorEvidence,
  );
  if (!evidence) {
    fail("invalid-request", "Embedded image history evidence is invalid.");
  }
  const expected = Uint8Array.from(input.expectedCurrentOperatorBytes);
  const expectedText = decodeLatin1(expected);
  const target = await resolveEmbeddedImageTransformTarget({
    sourceBytes: input.sourceBytes,
    pageNumber: evidence.pageNumber,
    imageOrder: evidence.imageOrder,
    withoutFrameClip: parsePdfNativeEmbeddedImageFrameClip(expectedText) === null,
  });
  if (!bytesEqual(target.current, expected)) {
    fail("precondition-mismatch", "Embedded image paint operator changed.");
  }
  // The revision that carried an alpha rewrote the page's Resources, so the
  // owner recorded before it is no longer the current one. Re-state the
  // object identities from the bytes being written now and keep only the
  // operator lineage from the prior evidence.
  const current: PdfNativeEmbeddedImageTransformEvidence = Object.freeze({
    ...evidence,
    page: Object.freeze({ ...target.graph.page.page }),
    resourceOwner: Object.freeze({
      ...nativeImageResourceOwner(target.graph, target.invocation.host),
    }),
    image: Object.freeze({ ...target.image }),
    contentStream: Object.freeze({ ...target.invocation.stream }),
  });
  const desired = Uint8Array.from(input.desiredOperatorBytes);
  // Re-applying a wrapper that names an alpha needs that resource present;
  // it already is, because the revision that first wrote it installed it.
  const desiredText = decodeLatin1(desired);
  const wrapper = parsePdfNativeEmbeddedImageTransformOperator(
    parsePdfNativeEmbeddedImageFrameClip(desiredText)?.drawText ?? desiredText,
  );
  if (wrapper?.graphicsStateName) {
    const states = await inspectMraiNativePageExtGStates(
      input.sourceBytes,
      evidence.pageNumber,
    );
    const entry = states.entries.find(
      (item) => item.resourceName === wrapper.graphicsStateName,
    );
    if (
      !entry ||
      !entry.supportedAppearanceOnly ||
      entry.fillOpacity === null ||
      (current.opacity !== undefined &&
        current.opacity !== null &&
        Math.abs(entry.fillOpacity - current.opacity) > 1e-9)
    ) {
      fail(
        "precondition-mismatch",
        "Embedded image alpha resource is no longer available.",
      );
    }
  }
  return buildNativeEmbeddedImageTransformPlan({
    target,
    imageOrder: evidence.imageOrder,
    evidence: current,
    desiredOperatorBytes: desired,
  });
}

export async function planMraiNativeEmbeddedImageReplacement(input: {
  readonly sourceBytes: Uint8Array;
  readonly selection: PdfNativeEmbeddedImageReplacementSelection;
  readonly replacementBytes: Uint8Array;
  readonly replacementMimeType: string;
}): Promise<PdfNativePreparedEmbeddedImageReplacement> {
  const pageNumber = exactPositiveInteger(
    input.selection.pageNumber,
    "embedded image page number",
  );
  const imageOrder = exactPositiveInteger(
    input.selection.imageOrder,
    "embedded image order",
    true,
  );
  const replacement = await inspectMraiNativeEmbeddedImageReplacement(
    input.replacementBytes,
    input.replacementMimeType,
  );
  return buildNativeEmbeddedImageReplacement({
    sourceBytes: input.sourceBytes,
    pageNumber,
    imageOrder,
    replacement,
    expectedCurrent: "original",
    desired: "replacement",
    expectedMatrix: input.selection.matrix ?? null,
  });
}

export async function planMraiNativeEmbeddedImageReplacementHistory(input: {
  readonly sourceBytes: Uint8Array;
  readonly expectedSourceSha256: string;
  readonly priorEvidence: PdfNativeEmbeddedImageReplacementEvidence;
  readonly expectedCurrent: "original" | "replacement";
  readonly desired: "original" | "replacement";
}): Promise<PdfNativePreparedEmbeddedImageReplacement> {
  if (
    (await sha256PdfBytes(input.sourceBytes)) !==
    exactSha256(input.expectedSourceSha256, "embedded image history source")
  ) {
    fail("source-mismatch", "Embedded image history source changed.");
  }
  const evidence = input.priorEvidence;
  if (
    evidence.schemaVersion !==
      MRAI_NATIVE_EMBEDDED_IMAGE_REPLACE_SCHEMA_VERSION ||
    !PDF_RESOURCE_NAME.test(evidence.originalResourceName) ||
    !PDF_RESOURCE_NAME.test(evidence.replacementResourceName) ||
    evidence.originalResourceName === evidence.replacementResourceName ||
    evidence.replacementImage.expectedObjectSha256 !==
      exactSha256(
        evidence.replacementImage.expectedObjectSha256,
        "replacement image object",
      )
  ) {
    fail("invalid-request", "Embedded image history evidence is invalid.");
  }
  return buildNativeEmbeddedImageReplacement({
    sourceBytes: input.sourceBytes,
    pageNumber: evidence.pageNumber,
    imageOrder: evidence.imageOrder,
    replacement: evidence.replacement,
    priorEvidence: evidence,
    expectedCurrent: input.expectedCurrent,
    desired: input.desired,
  });
}

export async function verifyMraiNativeEmbeddedImageReplacementCandidate(
  sourceBytes: Uint8Array,
  prepared: PdfNativePreparedEmbeddedImageReplacement,
): Promise<PdfNativeReopenVerification> {
  if (
    prepared.sourceLength !== sourceBytes.length ||
    !bytesEqual(
      prepared.candidateBytes.slice(0, sourceBytes.length),
      sourceBytes,
    ) ||
    (await sha256PdfBytes(prepared.candidateBytes)) !== prepared.candidateSha256
  ) {
    fail(
      "verification-failed",
      "Embedded image candidate source custody failed.",
    );
  }
  const reopened = parsePdf(prepared.candidateBytes, true);
  if (
    reopened.current.offset !== prepared.incrementalXrefOffset ||
    reopened.current.previousOffset !== prepared.sourceXrefOffset ||
    reopened.previous?.offset !== prepared.sourceXrefOffset
  ) {
    fail("verification-failed", "Embedded image xref linkage is invalid.");
  }
  for (const item of prepared.revisedObjects) {
    const entry = reopened.current.entries.get(item.objectNumber);
    if (
      !entry?.inUse ||
      entry.generation !== item.generation ||
      entry.offset !== item.offset
    ) {
      fail(
        "verification-failed",
        "Embedded image object ownership is invalid.",
      );
    }
    const object = parseIndirectObject(
      prepared.candidateBytes,
      reopened,
      item.objectNumber,
      item.generation,
    );
    if ((await sha256PdfBytes(object.bytes)) !== item.sha256) {
      fail(
        "verification-failed",
        "Embedded image object failed reopen verification.",
      );
    }
  }
  return Object.freeze({
    verified: true,
    identity: Object.freeze({
      sourceSha256: await sha256PdfBytes(sourceBytes),
      ...prepared.evidence.contentStream,
    }),
    sourceXrefOffset: prepared.sourceXrefOffset,
    incrementalXrefOffset: prepared.incrementalXrefOffset,
    candidateSha256: prepared.candidateSha256,
  });
}

function replaceTrailerSizeAndAppendPrev(
  sourceTrailerRaw: string,
  size: number,
  previousOffset: number,
): string {
  // /XRefStm belongs to the hybrid section it was written in; the appended
  // classic section reaches those objects through /Prev instead.
  const trailerRaw = sourceTrailerRaw.replace(/\s*\/XRefStm\s+\d+\b/g, "");
  const sizes = [...trailerRaw.matchAll(/\/Size\s+(\d+)\b/g)];
  const previous = [...trailerRaw.matchAll(/\/Prev\s+(\d+)\b/g)];
  if (sizes.length !== 1 || previous.length > 1) {
    fail(
      "unsupported-document",
      "Source trailer cannot be incrementally linked.",
    );
  }
  const relinked =
    previous.length === 1
      ? (() => {
          const match = previous[0];
          const start = match.index! + match[0].lastIndexOf(match[1]);
          return (
            trailerRaw.slice(0, start) +
            String(previousOffset) +
            trailerRaw.slice(start + match[1].length)
          );
        })()
      : trailerRaw;
  const sizeMatch = [...relinked.matchAll(/\/Size\s+(\d+)\b/g)][0];
  const numberStart = sizeMatch.index! + sizeMatch[0].lastIndexOf(sizeMatch[1]);
  const resized =
    relinked.slice(0, numberStart) +
    String(size) +
    relinked.slice(numberStart + sizeMatch[1].length);
  const end = resized.lastIndexOf(">>");
  if (end < 2) fail("unsupported-document", "Source trailer is malformed.");
  return previous.length === 1
    ? resized
    : `${resized.slice(0, end).trimEnd()}\n/Prev ${previousOffset}\n>>`;
}

/** The /Info and /ID an xref-stream source's newest stream dictionary holds,
 * as written there (leading space included, or empty). A revision appended
 * over that source must keep them: omitting /Info makes every metadata field
 * disappear to conforming readers after an otherwise unrelated native edit.
 * Copied byte for byte, since an /ID string may hold raw digest bytes. */
function streamDocumentAuthority(sourceXref: PdfXrefSection): string {
  const preserved = ["Info", "ID"].flatMap((key) => {
    const matches = topLevelDictionaryEntries(sourceXref.trailerRaw).filter(
      (entry) => entry.key === key,
    );
    if (matches.length > 1) {
      fail("unsupported-document", `PDF xref ${key} entry is ambiguous.`);
    }
    return matches.length === 1
      ? [
          sourceXref.trailerRaw.slice(
            matches[0].keyStart,
            matches[0].valueEnd,
          ),
        ]
      : [];
  });
  return preserved.length > 0 ? ` ${preserved.join(" ")}` : "";
}

interface IncrementalXrefRow {
  readonly objectNumber: number;
  readonly generation: number;
  readonly offset: number;
}

/** The rows an appended section carries and the /Size a classic table over
 * them declares: the objects this revision writes, and every object the
 * source's newest entries miss (`relocated`) at the place the file holds it,
 * so a reader that never repairs reads the revision the way this engine read
 * its source. A written object's own row always stands for its number. */
function incrementalXrefRows(
  sourceXref: PdfXrefSection,
  written: readonly IncrementalXrefRow[],
): Readonly<{ rows: readonly IncrementalXrefRow[]; size: number }> {
  const numbers = new Set(written.map((row) => row.objectNumber));
  const ordered = [
    ...written,
    ...[...(sourceXref.relocated?.values() ?? [])].filter(
      (entry) => !numbers.has(entry.objectNumber),
    ),
  ].sort(
    (left, right) =>
      left.objectNumber - right.objectNumber ||
      left.generation - right.generation,
  );
  return Object.freeze({
    rows: ordered,
    size: Math.max(sourceXref.size, ...ordered.map((row) => row.objectNumber + 1)),
  });
}

/** Ordered rows cut into runs of consecutive object numbers: the subsections
 * of a classic table, and the /Index pairs of a cross-reference stream. */
function incrementalXrefRuns(
  rows: readonly IncrementalXrefRow[],
): readonly (readonly IncrementalXrefRow[])[] {
  const runs: IncrementalXrefRow[][] = [];
  for (const row of rows) {
    const run = runs.at(-1);
    if (run && run.at(-1)!.objectNumber + 1 === row.objectNumber) run.push(row);
    else runs.push([row]);
  }
  return runs;
}

/** The cross-reference section a revision appends at `sectionOffset`, through
 * its `%%EOF`. Over a classic or hybrid source it is a classic table whose
 * trailer is the source's own, relinked. Over a source whose newest section
 * is a cross-reference stream it is a cross-reference stream too: Apple's
 * reader (Preview, PDFKit) does not follow a classic trailer's /Prev into a
 * stream section and refuses the whole file ("failed to find start of
 * cross-reference section"), while a stream section linked to a stream
 * section opens everywhere. */
function incrementalXrefSection(
  sourceXref: PdfXrefSection,
  written: readonly IncrementalXrefRow[],
  sectionOffset: number,
): Uint8Array {
  const { rows, size } = incrementalXrefRows(sourceXref, written);
  if (sourceXref.kind === "stream") {
    return incrementalXrefStream(sourceXref, rows, size, sectionOffset);
  }
  const lines: string[] = [];
  for (const run of incrementalXrefRuns(rows)) {
    lines.push(`${run[0].objectNumber} ${run.length}`);
    for (const row of run) {
      lines.push(
        `${String(row.offset).padStart(10, "0")} ${String(row.generation).padStart(5, "0")} n `,
      );
    }
  }
  const trailer = replaceTrailerSizeAndAppendPrev(
    sourceXref.trailerRaw,
    size,
    sourceXref.offset,
  );
  // The trailer is copied from the source: its /ID strings hold raw bytes
  // (Distiller writes the digest unescaped), so it goes back byte for byte.
  return latin1Bytes(
    `xref\n${lines.join("\n")}\ntrailer\n${trailer}\nstartxref\n${sectionOffset}\n%%EOF\n`,
  );
}

/** One cross-reference stream object, numbered with the /Size a classic table
 * over the same rows would declare (so it is past every object the source and
 * this revision name), listing those rows and itself. It is written
 * unfiltered with the narrowest offset field that holds every offset: the
 * replay derives the same bytes with nothing left to an encoder's choices,
 * and every reader takes an unfiltered stream. */
function incrementalXrefStream(
  sourceXref: PdfXrefSection,
  rows: readonly IncrementalXrefRow[],
  objectNumber: number,
  sectionOffset: number,
): Uint8Array {
  if (objectNumber + 1 > 9_999_999_999 || sourceXref.entries.has(objectNumber)) {
    fail("unsupported-document", "Incremental xref stream object number is unavailable.");
  }
  const listed = [...rows, { objectNumber, generation: 0, offset: sectionOffset }];
  const largest = Math.max(...listed.map((row) => row.offset));
  let width = 1;
  while (largest >= 256 ** width) width += 1;
  const rowWidth = 1 + width + 2;
  const data = new Uint8Array(listed.length * rowWidth);
  listed.forEach((row, index) => {
    const start = index * rowWidth;
    data[start] = 1;
    let offset = row.offset;
    for (let place = width; place >= 1; place -= 1) {
      data[start + place] = offset % 256;
      offset = Math.floor(offset / 256);
    }
    data[start + width + 1] = row.generation >>> 8;
    data[start + width + 2] = row.generation & 0xff;
  });
  const index = incrementalXrefRuns(listed)
    .map((run) => `${run[0].objectNumber} ${run.length}`)
    .join(" ");
  const dictionary =
    `<< /Type /XRef /Size ${objectNumber + 1} /Root ${sourceXref.rootObjectNumber} ${sourceXref.rootGeneration} R` +
    `${streamDocumentAuthority(sourceXref)} /Prev ${sourceXref.offset}` +
    ` /Index [${index}] /W [1 ${width} 2] /Length ${data.length} >>`;
  return concatBytes([
    latin1Bytes(`${objectNumber} 0 obj\n${dictionary}\nstream\n`),
    data,
    ascii(`\nendstream\nendobj\nstartxref\n${sectionOffset}\n%%EOF\n`),
  ]);
}

function makeIncrementalRevision(
  sourceBytes: Uint8Array,
  sourceXref: PdfXrefSection,
  objectNumber: number,
  generation: number,
  revisedObject: Uint8Array,
): {
  readonly candidateBytes: Uint8Array;
  readonly changedObjectOffset: number;
  readonly incrementalXrefOffset: number;
} {
  const separator =
    sourceBytes.at(-1) === 0x0a || sourceBytes.at(-1) === 0x0d
      ? new Uint8Array()
      : ascii("\n");
  const changedObjectOffset = sourceBytes.length + separator.length;
  if (changedObjectOffset > 9_999_999_999) {
    fail("unsupported-document", "Incremental object offset is too large.");
  }
  const objectBoundary = concatBytes([revisedObject, ascii("\n")]);
  const incrementalXrefOffset = changedObjectOffset + objectBoundary.length;
  if (incrementalXrefOffset > 9_999_999_999) {
    fail("unsupported-document", "Incremental xref offset is too large.");
  }
  const xref = incrementalXrefSection(
    sourceXref,
    [{ objectNumber, generation, offset: changedObjectOffset }],
    incrementalXrefOffset,
  );
  return Object.freeze({
    candidateBytes: concatBytes([sourceBytes, separator, objectBoundary, xref]),
    changedObjectOffset,
    incrementalXrefOffset,
  });
}

interface IncrementalObjectRevision {
  readonly objectNumber: number;
  readonly generation: number;
  readonly bytes: Uint8Array;
}

function escapeRegex(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/** Whether a paragraph rewrite's block draws with an installed face exactly
 * the glyphs its package holds: every string it shows while the face is set
 * (from its `Tf` to the next) is codes of the package, and every glyph of the
 * package is drawn at least once. A block sets the face as often as the text
 * holds characters only it has, so the codes repeat and interleave with the
 * paragraph's own fonts; the rewrite's own proof places every one of them. */
/** The parts of a run a property edit writes in an installed face (a part
 * of a line set in Arial, WPS): each selection of the face is followed by
 * that part's state and one `<codes> Tj` before anything selects another
 * font, and the codes of all of them, in order, are the package's. */
function pdfNativePropertyPartsUseFontPackage(
  replacement: string,
  fontPackage: Readonly<{ resourceName: string; encodedHex: string }>,
): boolean {
  const selects = new RegExp(`/${escapeRegex(fontPackage.resourceName)}\\s+[-+]?(?:\\d+(?:\\.\\d*)?|\\.\\d+)\\s+Tf\\b`, "gu");
  let codes = "";
  let parts = 0;
  for (const select of replacement.matchAll(selects)) {
    const from = select.index! + select[0].length;
    const next = replacement.slice(from).search(/\bTf\b/u);
    const run = next < 0 ? replacement.slice(from) : replacement.slice(from, from + next);
    const shows = [...run.matchAll(/<([0-9A-F]*)>\s*Tj/gu)];
    if (shows.length !== 1 || /\bTJ\b|'|"/u.test(run)) return false;
    codes += shows[0][1];
    parts += 1;
  }
  return parts > 0 && codes === fontPackage.encodedHex;
}

function pdfNativeParagraphBlockUsesFontPackage(
  replacement: string,
  fontPackage: Readonly<{ resourceName: string; encodedHex: string }>,
): boolean {
  if (fontPackage.encodedHex.length < 4 || fontPackage.encodedHex.length % 4 !== 0) return false;
  const packaged = new Set(fontPackage.encodedHex.match(/.{4}/gu) ?? []);
  const drawn = new Set<string>();
  const selects = new RegExp(`/${escapeRegex(fontPackage.resourceName)}\\s+[-+]?(?:\\d+(?:\\.\\d*)?|\\.\\d+)\\s+Tf\\b`, "gu");
  for (const select of replacement.matchAll(selects)) {
    const from = select.index! + select[0].length;
    const next = replacement.slice(from).search(/\bTf\b/u);
    const run = next < 0 ? replacement.slice(from) : replacement.slice(from, from + next);
    // Only whole arrays shown by TJ; any other show of the face is not the
    // rewrite's block.
    if (/\bTj\b|'|"/u.test(run)) return false;
    for (const array of run.matchAll(/\[([^\]]*)\]\s*TJ/gu)) {
      for (const hex of array[1].matchAll(/<([0-9A-F]*)>/gu)) {
        if (hex[1].length % 4 !== 0) return false;
        for (const code of hex[1].match(/.{4}/gu) ?? []) {
          if (!packaged.has(code)) return false;
          drawn.add(code);
        }
      }
    }
  }
  return drawn.size === packaged.size;
}

/** An object's bytes as an object of its own: a member of an object stream
 * is only its value there, and is written out under its own header (mirror
 * of the server's _standalone_object_raw). */
function standaloneObjectBytes(object: ParsedIndirectObject): Uint8Array {
  if (object.storage === "direct") return object.bytes;
  return concatBytes([
    ascii(`${object.objectNumber} ${object.generation} obj\n`),
    object.bytes,
    ascii("\nendobj"),
  ]);
}

/** A content stream's decoded bytes written back under the encoding the
 * producer chose (mirror of the server's _encode_like_target). */
function encodeLikeTarget(
  encoding: "identity" | "flate" | "ascii85-flate" | "ascii-hex" | "lzw",
  decoded: Uint8Array,
): Uint8Array {
  return encoding === "flate"
    ? encodeDeterministicFlateStream(decoded)
    : encoding === "ascii85-flate"
      ? encodeAscii85Stream(encodeDeterministicFlateStream(decoded))
      : encoding === "ascii-hex"
        ? encodeAsciiHexStream(decoded)
        : encoding === "lzw"
          ? encodeLzwStream(decoded)
          : decoded;
}

/** The stream object that draws a Form more than once, with the one drawing
 * an isolation edits renamed: the name operand at `operandStart` must be the
 * Form's own name (written without escapes) consumed by a `Do` (mirror of the
 * server's _renamed_drawing_stream). */
function renamedDrawingStreamBytes(
  sourceBytes: Uint8Array,
  stream: ParsedIndirectObject,
  resourceName: string,
  instance: PdfNativeFormInstanceIsolationInstance,
): Uint8Array {
  assertTargetStream(stream);
  const decoded = stream.streamBytes;
  const invocations = scanNativeImageInvocations(decoded, [
    { stream: { objectNumber: stream.objectNumber, generation: stream.generation }, start: 0, end: decoded.length },
  ]);
  const named = `/${resourceName}`;
  const tokenEnd = instance.operandStart + named.length;
  if (
    !invocations.some(
      (invocation) =>
        invocation.resourceName === resourceName &&
        invocation.operatorSpan.start === instance.operandStart,
    ) ||
    asLatin1(decoded.subarray(instance.operandStart, tokenEnd)) !== named ||
    (tokenEnd < decoded.length &&
      !isPdfNameDelimiter(String.fromCharCode(decoded[tokenEnd])))
  ) {
    fail("precondition-mismatch", "Renamed Form drawing does not match the stream.");
  }
  const renamed = concatBytes([
    decoded.subarray(0, instance.operandStart),
    ascii(`/${instance.resourceName}`),
    decoded.subarray(tokenEnd),
  ]);
  return rewriteStreamObjectBytes(
    sourceBytes,
    stream,
    encodeLikeTarget(stream.streamEncoding, renamed),
  );
}

function isPdfNameDelimiter(character: string): boolean {
  return /[\x00\t\n\f\r ()<>\[\]{}\/%]/.test(character);
}

/** Add one name to the direct /XObject dictionary an owner holds (under its
 * /Resources, or at the top of a resources dictionary), just before that
 * dictionary closes; the name must be new (mirror of the server's
 * _add_xobject_mapping). */
function addDirectXObjectMapping(
  objectBytes: Uint8Array,
  resourceName: string,
  target: Readonly<{ objectNumber: number; generation: number }>,
  ownerIsResourceDictionary: boolean,
): Uint8Array {
  const source = asLatin1(objectBytes);
  const header = /^(\d+)\s+(\d+)\s+obj\b/.exec(source);
  if (!header) {
    fail("unsupported-document", "Resource owner is not a direct object.");
  }
  const topStart = source.indexOf("<<", header[0].length);
  if (topStart < 0) {
    fail("unsupported-document", "Resource owner dictionary is missing.");
  }
  const topEnd = dictionaryEnd(source, topStart);
  const resources = ownerIsResourceDictionary
    ? Object.freeze({ start: topStart, end: topEnd })
    : directSubdictionarySpan(source, "Resources", topStart, topEnd, "Form instance Resources");
  const xobjects = directSubdictionarySpan(
    source,
    "XObject",
    resources.start,
    resources.end,
    "Form instance XObject resources",
  );
  if (
    topLevelDictionaryEntries(source.slice(xobjects.start, xobjects.end)).some(
      (entry) => entry.key === resourceName,
    )
  ) {
    fail("unsupported-document", "Form instance resource name is already taken.");
  }
  const insertAt = xobjects.end - 2;
  return concatBytes([
    objectBytes.slice(0, insertAt),
    ascii(` /${resourceName} ${target.objectNumber} ${target.generation} R`),
    objectBytes.slice(insertAt),
  ]);
}

function replaceIndirectObjectHeader(
  objectBytes: Uint8Array,
  objectNumber: number,
  generation: number,
): Uint8Array {
  const source = asLatin1(objectBytes);
  const header = /^(\d+)\s+(\d+)\s+obj\b/.exec(source);
  if (!header) {
    fail("unsupported-document", "Form indirect object header is malformed.");
  }
  return concatBytes([
    ascii(`${objectNumber} ${generation} obj`),
    objectBytes.slice(header[0].length),
  ]);
}

function directSubdictionarySpan(
  source: string,
  key: string,
  start: number,
  end: number,
  label: string,
): Readonly<{ start: number; end: number }> {
  const region = source.slice(start, end);
  if (new RegExp(`/${escapeRegex(key)}\\s+\\d+\\s+\\d+\\s+R\\b`).test(region)) {
    fail("unsupported-document", `Indirect ${label} are preserve-only.`);
  }
  const matches = [
    ...region.matchAll(new RegExp(`/${escapeRegex(key)}\\s*<<`, "g")),
  ];
  if (matches.length !== 1) {
    fail("unsupported-document", `${label} are missing or ambiguous.`);
  }
  const dictionaryStart =
    start + matches[0].index! + matches[0][0].lastIndexOf("<<");
  return Object.freeze({
    start: dictionaryStart,
    end: dictionaryEnd(source, dictionaryStart),
  });
}

/** Rewrite the one reference by which a parent names this link: a resources
 * dictionary under the parent's own /Resources, or a Form under an /XObject
 * resource name. */
function replaceIsolationLink(
  objectBytes: Uint8Array,
  link: Readonly<{ resourceName: string; kind?: "xobject" | "resources" }>,
  previous: Readonly<{ objectNumber: number; generation: number }>,
  next: Readonly<{ objectNumber: number; generation: number }>,
  ownerIsResourceDictionary = false,
): Uint8Array {
  if (link.kind !== "resources") {
    return replaceDirectXObjectMapping(
      objectBytes,
      link.resourceName,
      previous,
      next,
      ownerIsResourceDictionary,
    );
  }
  const source = asLatin1(objectBytes);
  const header = /^(\d+)\s+(\d+)\s+obj\b/.exec(source);
  if (!header) {
    fail("unsupported-document", "Resource owner is not a direct object.");
  }
  const topStart = source.indexOf("<<", header[0].length);
  if (topStart < 0) {
    fail("unsupported-document", "Resource owner dictionary is missing.");
  }
  const topEnd = dictionaryEnd(source, topStart);
  const entry = oneTopLevelDictionaryEntry(
    source.slice(topStart, topEnd),
    link.resourceName,
    `Page ${link.resourceName}`,
  );
  if (!entry) {
    fail("unsupported-document", "Selected page resource reference is missing or ambiguous.");
  }
  const value = source.slice(topStart + entry.valueStart, topStart + entry.valueEnd).trim();
  if (value !== `${previous.objectNumber} ${previous.generation} R`) {
    fail("unsupported-document", "Selected page resource reference is missing or ambiguous.");
  }
  return latin1Bytes(
    `${source.slice(0, topStart + entry.valueStart)}${next.objectNumber} ${next.generation} R${source.slice(topStart + entry.valueEnd)}`,
  );
}

function replaceDirectXObjectMapping(
  objectBytes: Uint8Array,
  resourceName: string,
  previous: Readonly<{ objectNumber: number; generation: number }>,
  next: Readonly<{ objectNumber: number; generation: number }>,
  ownerIsResourceDictionary = false,
): Uint8Array {
  const source = asLatin1(objectBytes);
  const header = /^(\d+)\s+(\d+)\s+obj\b/.exec(source);
  if (!header) {
    fail("unsupported-document", "Resource owner is not a direct object.");
  }
  const topStart = source.indexOf("<<", header[0].length);
  if (topStart < 0) {
    fail("unsupported-document", "Resource owner dictionary is missing.");
  }
  const topEnd = dictionaryEnd(source, topStart);
  // An owner reached through a /Resources reference is that dictionary: its
  // /XObject sits at the top level instead of under a /Resources key.
  const resources = ownerIsResourceDictionary
    ? Object.freeze({ start: topStart, end: topEnd })
    : directSubdictionarySpan(
        source,
        "Resources",
        topStart,
        topEnd,
        "Form instance Resources",
      );
  const xobjects = directSubdictionarySpan(
    source,
    "XObject",
    resources.start,
    resources.end,
    "Form instance XObject resources",
  );
  const mapping = new RegExp(
    `/${escapeRegex(resourceName)}\\s+${previous.objectNumber}\\s+${previous.generation}\\s+R\\b`,
    "g",
  );
  const matches = [
    ...source.slice(xobjects.start, xobjects.end).matchAll(mapping),
  ];
  if (matches.length !== 1) {
    fail(
      "unsupported-document",
      "Selected Form resource mapping is missing or ambiguous.",
    );
  }
  const matchStart = xobjects.start + matches[0].index!;
  const matchEnd = matchStart + matches[0][0].length;
  const replacement = `/${resourceName} ${next.objectNumber} ${next.generation} R`;
  return concatBytes([
    objectBytes.slice(0, matchStart),
    ascii(replacement),
    objectBytes.slice(matchEnd),
  ]);
}









export function appendMraiNativeObjectRevisions(
  sourceBytes: Uint8Array,
  revisions: readonly Readonly<{ objectNumber: number; generation: number; bytes: Uint8Array }>[],
): Uint8Array {
  const parsed = parsePdf(sourceBytes, false);
  return makeMultiObjectIncrementalRevision(sourceBytes, parsed.current, revisions,
    { min: 1, max: 256, label: "Appended object" }).candidateBytes;
}













export function squashMraiNativeRevisions(
  base: Uint8Array,
  candidate: Uint8Array,
): Uint8Array | null {
  if (candidate.length <= base.length) return null;
  for (let index = 0; index < base.length; index += 1) {
    if (candidate[index] !== base[index]) return null;
  }
  const parsedBase = parsePdf(base, false);
  const parsedCandidate = parsePdf(candidate, false);
  const appended = parsedCandidate.sections.length - parsedBase.sections.length;
  if (appended < 2) return null;
  if (parsedCandidate.sections[appended].offset !== parsedBase.current.offset) return null;
  const source = decodedSourceLatin1(candidate);
  const newest = new Map<number, XrefEntry>();
  for (const section of parsedCandidate.sections.slice(0, appended)) {
    // A cross-reference stream section lists itself under the last number.
    const self = section.kind === "stream" ? section.size - 1 : null;
    for (const entry of section.entries.values()) {
      if (entry.objectNumber === self || newest.has(entry.objectNumber)) continue;
      if (!entry.inUse || entry.kind !== "direct" || entry.offset < base.length) return null;
      newest.set(entry.objectNumber, entry);
    }
  }
  if (newest.size === 0 || newest.size > 256) return null;
  const revisions = [...newest.values()]
    .sort((left, right) => left.objectNumber - right.objectNumber)
    .map((entry) => {
      const object = parseDirectObjectAtOffset(candidate, entry.offset,
        { objectNumber: entry.objectNumber, generation: entry.generation }, source,
        (reference) => resolveStreamLengthReference(candidate, parsedCandidate, reference, source), "raw");
      return Object.freeze({ objectNumber: entry.objectNumber, generation: entry.generation,
        bytes: candidate.slice(object.start, object.end) });
    });
  return makeMultiObjectIncrementalRevision(base, parsedBase.current, revisions,
    { min: 1, max: 256, label: "Squashed" }).candidateBytes;
}

function makeMultiObjectIncrementalRevision(
  sourceBytes: Uint8Array,
  sourceXref: PdfXrefSection,
  revisions: readonly IncrementalObjectRevision[],
  limits: Readonly<{ min: number; max: number; label: string }> = {
    min: 2,
    max: 16,
    label: "Form isolation",
  },
): Readonly<{
  candidateBytes: Uint8Array;
  incrementalXrefOffset: number;
  offsets: ReadonlyMap<string, number>;
}> {
  if (revisions.length < limits.min || revisions.length > limits.max) {
    fail("invalid-request", `${limits.label} revision set is invalid.`);
  }
  const seen = new Set<string>();
  const separator =
    sourceBytes.at(-1) === 0x0a || sourceBytes.at(-1) === 0x0d
      ? new Uint8Array()
      : ascii("\n");
  const parts: Uint8Array[] = [sourceBytes, separator];
  const offsets = new Map<string, number>();
  let cursor = sourceBytes.length + separator.length;
  for (const revision of revisions) {
    assertObjectReference(revision.objectNumber, revision.generation);
    const key = `${revision.objectNumber}:${revision.generation}`;
    if (seen.has(key)) {
      fail("invalid-request", "Incremental object ownership is duplicated.");
    }
    seen.add(key);
    if (cursor > 9_999_999_999) {
      fail("unsupported-document", "Incremental object offset is too large.");
    }
    offsets.set(key, cursor);
    parts.push(revision.bytes, ascii("\n"));
    cursor += revision.bytes.length + 1;
  }
  const incrementalXrefOffset = cursor;
  parts.push(
    incrementalXrefSection(
      sourceXref,
      revisions.map((revision) => ({
        objectNumber: revision.objectNumber,
        generation: revision.generation,
        offset: offsets.get(`${revision.objectNumber}:${revision.generation}`)!,
      })),
      incrementalXrefOffset,
    ),
  );
  return Object.freeze({
    candidateBytes: concatBytes(parts),
    incrementalXrefOffset,
    offsets,
  });
}

const MANAGED_OUTLINE_ID = /^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$/;

function outlineRefEqual(
  left: PdfNativeOutlineObjectRef | null,
  right: PdfNativeOutlineObjectRef | null,
): boolean {
  return (
    left === right ||
    (left !== null &&
      right !== null &&
      left.objectNumber === right.objectNumber &&
      left.generation === right.generation)
  );
}

function optionalOutlineReference(
  source: string,
  key: string,
  label: string,
): PdfNativeOutlineObjectRef | null {
  return oneTopLevelReference(source, key, label, true);
}

function signedOutlineCount(
  source: string,
  allowMissing: boolean,
): number | null {
  const entry = oneTopLevelDictionaryEntry(
    source,
    "Count",
    "PDF outline Count",
    allowMissing,
  );
  if (!entry) return null;
  const raw = source.slice(entry.valueStart, entry.valueEnd);
  if (!/^-?\d+$/.test(raw)) {
    fail("unsupported-document", "PDF outline Count is malformed.");
  }
  const count = Number(raw);
  if (!Number.isSafeInteger(count) || Math.abs(count) > 100_000) {
    fail("unsupported-document", "PDF outline Count is unsupported.");
  }
  return count;
}

function outlineOwnershipMarker(
  source: string,
  key: string,
): true | false | "/Root" | null {
  const entry = oneTopLevelDictionaryEntry(
    source,
    key,
    `PDF outline ${key}`,
    true,
  );
  if (!entry) return null;
  const value = source.slice(entry.valueStart, entry.valueEnd);
  if (value === "true") return true;
  if (value === "false") return false;
  if (value === "/Root") return "/Root";
  fail("unsupported-document", `PDF outline ${key} is malformed.`);
}

function directOutlineName(
  source: string,
  key: string,
): string | null {
  const entry = oneTopLevelDictionaryEntry(
    source,
    key,
    `PDF outline ${key}`,
    true,
  );
  if (!entry) return null;
  if (source[entry.valueStart] !== "/") {
    fail("unsupported-document", `PDF outline ${key} is malformed.`);
  }
  const end = pdfNameEnd(source, entry.valueStart, entry.valueEnd);
  if (end !== entry.valueEnd) {
    fail("unsupported-document", `PDF outline ${key} is malformed.`);
  }
  return decodePdfName(source, entry.valueStart, end);
}

function managedOutlineId(source: string): string | null {
  const entry = oneTopLevelDictionaryEntry(
    source,
    "MRAIId",
    "MRAI managed outline id",
    true,
  );
  if (!entry) return null;
  const raw = source.slice(entry.valueStart, entry.valueEnd);
  const match = /^\(([A-Za-z0-9][A-Za-z0-9._:-]{0,127})\)$/.exec(raw);
  if (!match) {
    fail("unsupported-document", "MRAI managed outline id is malformed.");
  }
  return match[1];
}

function outlineStringEnd(
  source: string,
  start: number,
  limit: number,
): number | null {
  if (source[start] === "(") return pdfLiteralEnd(source, start, limit);
  if (source[start] === "<" && source[start + 1] !== "<") {
    return pdfHexEnd(source, start, limit);
  }
  return null;
}

/**
 * A source node's /Title is preserve-only, but it still has to be one PDF
 * string the server can read back, and it may be written as an indirect
 * string object: hyperref/pdfTeX does that for every bookmark it emits.
 * Resolve that form here so both sides admit exactly the same documents.
 */
function safeOutlineTitle(
  bytes: Uint8Array,
  parsed: ParsedPdf,
  source: string,
): void {
  const entry = oneTopLevelDictionaryEntry(
    source,
    "Title",
    "PDF outline Title",
    true,
  );
  if (!entry) fail("unsupported-document", "PDF outline Title is missing.");
  const raw = source.slice(entry.valueStart, entry.valueEnd);
  const direct = outlineStringEnd(source, entry.valueStart, entry.valueEnd);
  if (direct !== null) {
    if (direct !== entry.valueEnd) {
      fail("unsupported-document", "PDF outline Title is malformed.");
    }
    return;
  }
  if (!/^\d/.test(raw)) {
    fail("unsupported-document", "PDF outline Title is not one PDF string.");
  }
  const reference = parseTopLevelReference(
    source,
    entry.valueStart,
    entry.valueEnd,
    "PDF outline Title",
  );
  if (reference.end !== entry.valueEnd) {
    fail("unsupported-document", "PDF outline Title is malformed.");
  }
  const object = parseIndirectObject(
    bytes,
    parsed,
    reference.reference.objectNumber,
    reference.reference.generation,
  );
  if (object.streamBytes !== null) {
    fail("unsupported-document", "Stream-backed PDF outline titles are preserve-only.");
  }
  const text = asLatin1(object.bytes);
  let start = 0;
  if (object.storage === "direct") {
    const header = /^\d+\s+\d+\s+obj\b/.exec(text);
    if (!header) {
      fail("unsupported-document", "PDF outline Title object is malformed.");
    }
    start = skipWhitespace(text, header[0].length);
  } else {
    start = skipWhitespace(text, 0);
  }
  const end = outlineStringEnd(text, start, text.length);
  if (end === null) {
    fail("unsupported-document", "PDF outline Title is not one PDF string.");
  }
  const tail = skipWhitespace(text, end);
  if (object.storage === "direct" ? text.slice(tail, tail + 6) !== "endobj" : tail !== text.length) {
    fail("unsupported-document", "PDF outline Title object is malformed.");
  }
}

const STANDARD_NAMED_OUTLINE_ACTIONS = new Set(["NextPage", "PrevPage", "FirstPage", "LastPage"]);

function safeOutlineAction(
  bytes: Uint8Array,
  parsed: ParsedPdf,
  source: string,
): void {
  const action = oneTopLevelDictionaryEntry(
    source,
    "A",
    "PDF outline action",
    true,
  );
  if (!action) return;
  const raw = source.slice(action.valueStart, action.valueEnd);
  const actionSource = raw.startsWith("<<")
    ? raw
    : (() => {
        const reference = parseTopLevelReference(
          source,
          action.valueStart,
          action.valueEnd,
          "PDF outline action",
        );
        if (reference.end !== action.valueEnd) {
          fail("unsupported-document", "PDF outline action is malformed.");
        }
        return asLatin1(
          parseIndirectObject(
            bytes,
            parsed,
            reference.reference.objectNumber,
            reference.reference.generation,
          ).bytes,
        );
      })();
  if (
    oneTopLevelDictionaryEntry(
      actionSource,
      "Next",
      "PDF outline action chain",
      true,
    )
  ) {
    fail(
      "unsupported-document",
      "Chained PDF outline actions are preserve-only.",
    );
  }
  const subtype = directOutlineName(actionSource, "S");
  if (subtype === "Named") {
    // The four page-navigation names the specification defines (12.6.4.11)
    // reach no script, file or URL; a bookmark may carry one unchanged.
    const name = directOutlineName(actionSource, "N");
    if (name === null || !STANDARD_NAMED_OUTLINE_ACTIONS.has(name)) {
      fail(
        "unsupported-document",
        "Only the standard page-navigation named outline actions can be preserved during mutation.",
      );
    }
    return;
  }
  if (subtype !== "GoTo") {
    fail(
      "unsupported-document",
      "Only in-document GoTo outline actions can be preserved during mutation.",
    );
  }
  const destination = oneTopLevelDictionaryEntry(
    actionSource,
    "D",
    "PDF GoTo outline action destination",
  );
  if (
    !destination ||
    !["(", "<", "/", "["].includes(actionSource[destination.valueStart]) ||
    actionSource.slice(destination.valueStart, destination.valueStart + 2) === "<<"
  ) {
    fail(
      "unsupported-document",
      "PDF GoTo outline action destination is missing or unsafe.",
    );
  }
}

function rewriteOutlineDictionaryObject(
  object: ParsedIndirectObject,
  changes: Readonly<Record<string, string | null>>,
): Uint8Array {
  if (object.streamBytes !== null) {
    fail("unsupported-document", "Stream-backed outline dictionaries are preserve-only.");
  }
  const original: Uint8Array = Uint8Array.from(object.bytes);
  const source = asLatin1(original);
  const entries = topLevelDictionaryEntries(source);
  const dictionaryStart = source.indexOf("<<");
  const dictionaryLimit = dictionaryEnd(source, dictionaryStart) - 2;
  const edits: Array<{ start: number; end: number; bytes: Uint8Array }> = [];
  const additions: string[] = [];
  for (const [key, value] of Object.entries(changes)) {
    const matches = entries.filter((entry) => entry.key === key);
    if (matches.length > 1) {
      fail("unsupported-document", `PDF outline ${key} is ambiguous.`);
    }
    if (matches.length === 0) {
      if (value !== null) additions.push(`\n/${key} ${value}`);
      continue;
    }
    const entry = matches[0];
    edits.push({
      start: value === null ? entry.keyStart : entry.valueStart,
      end: entry.valueEnd,
      bytes: value === null ? new Uint8Array() : ascii(value),
    });
  }
  if (additions.length > 0) {
    let insertionStart = dictionaryLimit;
    while (insertionStart > dictionaryStart + 2 && /\s/.test(source[insertionStart - 1])) {
      insertionStart -= 1;
    }
    edits.push({
      start: insertionStart,
      end: dictionaryLimit,
      bytes: ascii(`${additions.join("")}\n`),
    });
  }
  edits.sort((left, right) => right.start - left.start);
  let revised = original;
  for (const edit of edits) {
    revised = concatBytes([
      revised.slice(0, edit.start),
      edit.bytes,
      revised.slice(edit.end),
    ]);
  }
  if (object.storage === "direct") return revised;
  // A compressed member carries only the object's own value, but the bytes
  // around it inside the object stream (a separating newline, say) belong to
  // the container. Re-emitting them would make this revision differ from the
  // server's replay, which rebuilds the object from its dictionary alone.
  const revisedSource = asLatin1(revised);
  const revisedStart = revisedSource.indexOf("<<");
  return concatBytes([
    ascii(`${object.objectNumber} ${object.generation} obj\n`),
    revised.slice(revisedStart, dictionaryEnd(revisedSource, revisedStart)),
    ascii("\nendobj"),
  ]);
}

function outlineUtf16Hex(value: string): string {
  let hex = "FEFF";
  for (let index = 0; index < value.length; index += 1) {
    hex += value.charCodeAt(index).toString(16).padStart(4, "0").toUpperCase();
  }
  return `<${hex}>`;
}

function outlineNumber(value: number | null): string {
  if (value === null) return "null";
  if (!Number.isFinite(value) || Math.abs(value) > 10_000_000) {
    fail("invalid-request", "Managed outline XYZ coordinate is invalid.");
  }
  if (Object.is(value, -0) || value === 0) return "0";
  return value.toFixed(8).replace(/0+$/, "").replace(/\.$/, "");
}

function serializeManagedOutlineNode(
  ref: PdfNativeOutlineObjectRef,
  parent: PdfNativeOutlineObjectRef,
  previous: PdfNativeOutlineObjectRef | null,
  next: PdfNativeOutlineObjectRef | null,
  node: PdfNativeManagedOutlineNode,
  childRefs: readonly PdfNativeOutlineObjectRef[],
): Uint8Array {
  const descendants = (item: PdfNativeManagedOutlineNode): number =>
    item.children.reduce((total, child) => total + 1 + descendants(child), 0);
  const parts = [
    `${ref.objectNumber} ${ref.generation} obj`,
    "<<",
    `/Title ${outlineUtf16Hex(node.title)}`,
    `/Parent ${parent.objectNumber} ${parent.generation} R`,
  ];
  if (previous) parts.push(`/Prev ${previous.objectNumber} ${previous.generation} R`);
  if (next) parts.push(`/Next ${next.objectNumber} ${next.generation} R`);
  if (childRefs.length > 0) {
    parts.push(`/First ${childRefs[0].objectNumber} ${childRefs[0].generation} R`);
    const last = childRefs.at(-1)!;
    parts.push(`/Last ${last.objectNumber} ${last.generation} R`);
    parts.push(`/Count ${descendants(node)}`);
  }
  parts.push(
    `/Dest [${node.destination.page.objectNumber} ${node.destination.page.generation} R /XYZ ${outlineNumber(node.destination.left)} ${outlineNumber(node.destination.top)} ${outlineNumber(node.destination.zoom)}]`,
    "/MRAIManaged true",
    `/MRAIId (${node.id})`,
  );
  parts.push(">>", "endobj");
  return ascii(parts.join("\n"));
}

interface InspectedOutlineTree {
  readonly rootRef: PdfNativeOutlineObjectRef | null;
  readonly rootObject: ParsedIndirectObject | null;
  readonly sourceFirst: PdfNativeOutlineObjectRef | null;
  readonly sourceLast: PdfNativeOutlineObjectRef | null;
  readonly sourceVisibleCount: number;
  readonly managedContainer: PdfNativeOutlineObjectRef | null;
  readonly managedRefsById: ReadonlyMap<string, PdfNativeOutlineObjectRef>;
  readonly managedIds: readonly string[];
}

function inspectOutlineTree(
  bytes: Uint8Array,
  parsed: ParsedPdf,
  catalogSource: string,
): InspectedOutlineTree {
  const rootRef = optionalOutlineReference(
    catalogSource,
    "Outlines",
    "PDF Catalog Outlines",
  );
  if (!rootRef) {
    return Object.freeze({
      rootRef: null,
      rootObject: null,
      sourceFirst: null,
      sourceLast: null,
      sourceVisibleCount: 0,
      managedContainer: null,
      managedRefsById: new Map(),
      managedIds: Object.freeze([]),
    });
  }
  const rootObject = parseIndirectObject(
    bytes,
    parsed,
    rootRef.objectNumber,
    rootRef.generation,
  );
  if (rootObject.streamBytes !== null) {
    fail("unsupported-document", "Stream-backed PDF Outlines are preserve-only.");
  }
  const rootSource = asLatin1(rootObject.bytes);
  const rootType = directOutlineName(rootSource, "Type");
  if (rootType !== null && rootType !== "Outlines") {
    fail("unsupported-document", "PDF Catalog Outlines has the wrong type.");
  }
  const first = optionalOutlineReference(rootSource, "First", "PDF outline First");
  const last = optionalOutlineReference(rootSource, "Last", "PDF outline Last");
  if ((first === null) !== (last === null)) {
    fail("unsupported-document", "PDF outline First/Last links are incomplete.");
  }
  const seen = new Set<string>();
  const managedRefsById = new Map<string, PdfNativeOutlineObjectRef>();
  const managedIds: string[] = [];
  let managedContainer: PdfNativeOutlineObjectRef | null = null;
  let sourceFirst: PdfNativeOutlineObjectRef | null = null;
  let sourceLast: PdfNativeOutlineObjectRef | null = null;
  let sourceVisibleCount = 0;

  const visitSiblings = (
    firstRef: PdfNativeOutlineObjectRef | null,
    lastRef: PdfNativeOutlineObjectRef | null,
    parent: PdfNativeOutlineObjectRef,
    depth: number,
    insideManaged: boolean,
  ): Readonly<{
    total: number;
    visibleWhenOpen: number;
    refs: readonly PdfNativeOutlineObjectRef[];
  }> => {
    if ((firstRef === null) !== (lastRef === null)) {
      fail("unsupported-document", "PDF outline child links are incomplete.");
    }
    if (!firstRef || !lastRef) {
      return Object.freeze({ total: 0, visibleWhenOpen: 0, refs: [] });
    }
    if (depth > 12) {
      fail("unsupported-document", "PDF outline depth exceeds 12.");
    }
    const refs: PdfNativeOutlineObjectRef[] = [];
    let total = 0;
    let visibleWhenOpen = 0;
    let previous: PdfNativeOutlineObjectRef | null = null;
    let current: PdfNativeOutlineObjectRef | null = firstRef;
    while (current) {
      const key = `${current.objectNumber}:${current.generation}`;
      if (
        seen.has(key) ||
        refs.length >= 4_096 ||
        seen.size >= 1_024 + 4_096
      ) {
        fail("unsupported-document", "PDF outline links are cyclic or too broad.");
      }
      seen.add(key);
      refs.push(current);
      const object = parseIndirectObject(
        bytes,
        parsed,
        current.objectNumber,
        current.generation,
      );
      if (object.streamBytes !== null) {
        fail("unsupported-document", "Stream-backed PDF outline nodes are preserve-only.");
      }
      const source = asLatin1(object.bytes);
      const declaredParent = optionalOutlineReference(source, "Parent", "PDF outline Parent");
      const declaredPrevious = optionalOutlineReference(source, "Prev", "PDF outline Prev");
      const declaredNext = optionalOutlineReference(source, "Next", "PDF outline Next");
      if (!outlineRefEqual(declaredParent, parent) || !outlineRefEqual(declaredPrevious, previous)) {
        fail("unsupported-document", "PDF outline Parent/Prev links are inconsistent.");
      }
      const marker = outlineOwnershipMarker(source, "MRAIManaged");
      const schema = directOutlineName(source, "MRAISchema");
      const id = managedOutlineId(source);
      const isContainer = marker === "/Root" && schema === "V1" && id === null;
      const isManagedNode = marker === true && schema === null && id !== null;
      if (
        marker === false ||
        (marker !== null && !isContainer && !isManagedNode)
      ) {
        fail("unsupported-document", "MRAI outline ownership marker is malformed.");
      }
      if (insideManaged !== isManagedNode) {
        if (!(depth === 0 && !insideManaged && isContainer)) {
          fail("unsupported-document", "Managed and source outline ownership is mixed.");
        }
      }
      if (isContainer) {
        if (managedContainer || declaredNext !== null) {
          fail("unsupported-document", "The MRAI outline subtree must be the final root child.");
        }
        managedContainer = current;
      } else if (isManagedNode) {
        if (managedRefsById.has(id!)) {
          fail("unsupported-document", "Managed outline ids are duplicated.");
        }
        managedRefsById.set(id!, current);
        managedIds.push(id!);
      } else {
        if (schema !== null || id !== null) {
          fail("unsupported-document", "Source outline contains reserved MRAI keys.");
        }
        safeOutlineTitle(bytes, parsed, source);
        safeOutlineAction(bytes, parsed, source);
        const dest = oneTopLevelDictionaryEntry(source, "Dest", "PDF outline Dest", true);
        const action = oneTopLevelDictionaryEntry(source, "A", "PDF outline action", true);
        if (dest && action) {
          fail("unsupported-document", "PDF outline destination ownership is ambiguous.");
        }
        if (depth === 0) {
          sourceFirst ??= current;
          sourceLast = current;
        }
      }
      const childFirst = optionalOutlineReference(source, "First", "PDF outline First");
      const childLast = optionalOutlineReference(source, "Last", "PDF outline Last");
      const children = visitSiblings(
        childFirst,
        childLast,
        current,
        isContainer ? 0 : depth + 1,
        insideManaged || isContainer || isManagedNode,
      );
      const count = signedOutlineCount(source, true);
      // No reader reads a /Count's magnitude -- TCPDF writes 0 on every node
      // it emits, and a producer that miscounts (py-pdf's mistitled-outlines
      // sample) draws the same tree -- only its sign, open or closed. A source
      // item's magnitude is therefore the source's own shape and is carried
      // as it is; the rows an open item reveals are counted from its subtree.
      // The items this writer manages it writes itself, so they still have
      // to count exactly.
      if (
        (isContainer || isManagedNode || insideManaged) &&
        count !== null &&
        count !== 0 &&
        Math.abs(count) !== children.visibleWhenOpen
      ) {
        fail("unsupported-document", "PDF outline Count does not match its subtree.");
      }
      if (!isContainer && !isManagedNode && depth === 0) {
        sourceVisibleCount +=
          1 + (count !== null && count > 0 ? children.visibleWhenOpen : 0);
      }
      visibleWhenOpen +=
        1 + (count !== null && count > 0 ? children.visibleWhenOpen : 0);
      total += 1 + children.total;
      previous = current;
      if (outlineRefEqual(current, lastRef)) {
        if (declaredNext !== null) {
          fail("unsupported-document", "PDF outline Last has a Next link.");
        }
        current = null;
      } else {
        if (!declaredNext) {
          fail("unsupported-document", "PDF outline sibling chain ends before Last.");
        }
        current = declaredNext;
      }
    }
    return Object.freeze({
      total,
      visibleWhenOpen,
      refs: Object.freeze(refs),
    });
  };

  visitSiblings(first, last, rootRef, 0, false);
  // Read for its syntax only; see below.
  signedOutlineCount(rootSource, true);
  // The root is never drawn as a row, so its /Count carries no open/closed
  // state and readers derive the tree from /First alone: TCPDF omits it and
  // pdfTeX writes it negative, and a producer that miscounts its items
  // miscounts the root with them. The source's value is its own; the
  // revision this planner appends always writes the spec-correct one.
  return Object.freeze({
    rootRef,
    rootObject,
    sourceFirst,
    sourceLast,
    sourceVisibleCount,
    managedContainer,
    managedRefsById,
    managedIds: Object.freeze(managedIds),
  });
}

/**
 * The page identities an outline destination may name, in reading order.
 * A bookmark never reads page content, resources or their streams, so this
 * mirrors the server's own outline page walk exactly: object type, declared
 * parent, /Kids and /Count. Reusing the full content-graph inspection here
 * refused legal documents over things a bookmark cannot touch — an empty page
 * carrying no /Contents, or a content stream whose /Length is indirect.
 */
function readOutlinePageReferences(
  bytes: Uint8Array,
  parsed: ParsedPdf,
  catalogSource: string,
): readonly PdfNativeOutlineObjectRef[] {
  const pagesRoot = oneTopLevelReference(
    catalogSource,
    "Pages",
    "PDF Catalog Pages reference",
  );
  if (!pagesRoot) {
    fail("unsupported-document", "PDF Catalog Pages reference is missing.");
  }
  const seen = new Set<string>();
  const leaves: PdfNativeOutlineObjectRef[] = [];
  const visit = (
    ref: PdfNativeOutlineObjectRef,
    parent: PdfNativeOutlineObjectRef | null,
    depth: number,
  ): number => {
    if (depth > 32 || seen.size > 100_000) {
      fail("unsupported-document", "PDF page tree is too deep or large.");
    }
    const key = `${ref.objectNumber}:${ref.generation}`;
    if (seen.has(key)) {
      fail("unsupported-document", "PDF page tree is cyclic or duplicated.");
    }
    seen.add(key);
    const object = parseIndirectObject(
      bytes,
      parsed,
      ref.objectNumber,
      ref.generation,
    );
    const source = asLatin1(object.bytes);
    const declaredParent = oneTopLevelReference(
      source,
      "Parent",
      "PDF page-tree Parent reference",
      true,
    );
    if (!outlineRefEqual(declaredParent, parent)) {
      fail("unsupported-document", "PDF page-tree Parent identity is invalid.");
    }
    const type = exactObjectType(source);
    if (type === "Page") {
      leaves.push(Object.freeze({ ...ref }));
      return 1;
    }
    if (type !== "Pages") {
      fail("unsupported-document", "PDF page-tree node type is invalid.");
    }
    const total = pageTreeKids(source).reduce(
      (count, kid) => count + visit(kid, ref, depth + 1),
      0,
    );
    if (total !== pageTreeCount(source)) {
      fail("unsupported-document", "PDF page-tree Count is inconsistent.");
    }
    return total;
  };
  visit(pagesRoot, null, 0);
  return Object.freeze(leaves);
}

/** A request's page map, checked against the file it is planned on: one
 * entry per page, each an upload page named once or null (an added page). */
function readOutlinePageMap(
  value: unknown,
  pageCount: number,
): readonly (number | null)[] {
  if (!Array.isArray(value) || value.length !== pageCount) {
    fail("invalid-request", "Managed outline page map does not match the file.");
  }
  const seen = new Set<number>();
  for (const entry of value) {
    if (entry === null) continue;
    if (!Number.isSafeInteger(entry) || (entry as number) < 1 || seen.has(entry as number)) {
      fail("invalid-request", "Managed outline page map is invalid.");
    }
    seen.add(entry as number);
  }
  return Object.freeze([...(value as (number | null)[])]);
}

/** The tree with each destination moved from the upload's page it names to
 * the file's page showing it; a page deleted from the file is refused. */
function managedOutlineOnFilePages(
  nodes: readonly PdfNativeOutlineMutationInputNode[],
  pageMap: readonly (number | null)[],
): readonly PdfNativeOutlineMutationInputNode[] {
  if (!Array.isArray(nodes)) fail("invalid-request", "Managed outline tree is invalid.");
  return nodes.map((node) => {
    if (!node || typeof node !== "object" || !node.destination || !Array.isArray(node.children)) {
      fail("invalid-request", "Managed outline tree is invalid.");
    }
    const index = Number.isSafeInteger(node.destination.pageNumber)
      ? pageMap.indexOf(node.destination.pageNumber)
      : -1;
    if (index < 0) fail("invalid-request", "Managed outline page is not in the file.");
    return {
      ...node,
      destination: { ...node.destination, pageNumber: index + 1 },
      children: managedOutlineOnFilePages(node.children, pageMap),
    };
  });
}

function normalizeManagedOutlineInput(
  nodes: readonly PdfNativeOutlineMutationInputNode[],
  pages: readonly PdfNativeOutlineObjectRef[],
): readonly PdfNativeManagedOutlineNode[] {
  if (!Array.isArray(nodes)) fail("invalid-request", "Managed outline tree is invalid.");
  const ids = new Set<string>();
  let count = 0;
  const walk = (
    items: readonly PdfNativeOutlineMutationInputNode[],
    depth: number,
  ): readonly PdfNativeManagedOutlineNode[] => {
    if (depth > 12) fail("invalid-request", "Managed outline depth exceeds 12.");
    return Object.freeze(
      items.map((node) => {
        count += 1;
        if (count > 1_024 || !node || typeof node !== "object") {
          fail("invalid-request", "Managed outline tree is too broad.");
        }
        if (!MANAGED_OUTLINE_ID.test(node.id) || ids.has(node.id)) {
          fail("invalid-request", "Managed outline id is invalid or duplicated.");
        }
        ids.add(node.id);
        if (
          typeof node.title !== "string" ||
          node.title.length < 1 ||
          node.title.length > OUTLINE_TITLE_MAX_LENGTH ||
          hasUnpairedUtf16Surrogate(node.title) ||
          node.title !== node.title.replace(/[\s\u0000-\u001f\u007f-\u009f]+/g, " ").trim()
        ) {
          fail("invalid-request", "Managed outline title is not normalized.");
        }
        const destination = node.destination;
        if (!destination || !Number.isSafeInteger(destination.pageNumber)) {
          fail("invalid-request", "Managed outline destination is invalid.");
        }
        const page = pages[destination.pageNumber - 1];
        if (!page) fail("invalid-request", "Managed outline page is out of range.");
        for (const value of [destination.left, destination.top]) {
          if (
            value !== null &&
            (typeof value !== "number" ||
              !Number.isFinite(value) ||
              Math.abs(value) > MANAGED_OUTLINE_MAX_ABS_COORDINATE)
          ) {
            fail("invalid-request", "Managed outline XYZ coordinate is invalid.");
          }
        }
        if (
          destination.zoom !== null &&
          (typeof destination.zoom !== "number" ||
            !Number.isFinite(destination.zoom) ||
            destination.zoom <= 0 ||
            destination.zoom > MANAGED_OUTLINE_MAX_ZOOM)
        ) {
          fail("invalid-request", "Managed outline zoom is invalid.");
        }
        if (!Array.isArray(node.children)) {
          fail("invalid-request", "Managed outline children are invalid.");
        }
        return Object.freeze({
          id: node.id,
          title: node.title,
          destination: Object.freeze({
            pageNumber: destination.pageNumber,
            page: Object.freeze({ ...page }),
            left: destination.left,
            top: destination.top,
            zoom: destination.zoom,
          }),
          children: walk(node.children, depth + 1),
        });
      }),
    );
  };
  return walk(nodes, 0);
}

/** Build one deterministic incremental revision for the complete MRAI subtree. */
interface SourceOutlineNode {
  readonly ref: PdfNativeOutlineObjectRef;
  readonly object: ParsedIndirectObject;
  count: number | null;
  readonly children: SourceOutlineNode[];
}

/** The file's own bookmarks below the outline root, MRAI's subtree left out
 * (`inspectOutlineTree` has already validated the whole tree). */
function readSourceOutlineTree(
  bytes: Uint8Array,
  parsed: ParsedPdf,
  inspected: InspectedOutlineTree,
): SourceOutlineNode[] {
  const build = (ref: PdfNativeOutlineObjectRef): SourceOutlineNode => {
    const object = parseIndirectObject(bytes, parsed, ref.objectNumber, ref.generation);
    const source = asLatin1(object.bytes);
    const children: SourceOutlineNode[] = [];
    let child = optionalOutlineReference(source, "First", "PDF outline First");
    while (child) {
      const node = build(child);
      children.push(node);
      child = optionalOutlineReference(asLatin1(node.object.bytes), "Next", "PDF outline Next");
    }
    return { ref, object, count: signedOutlineCount(source, true), children };
  };
  const top: SourceOutlineNode[] = [];
  let current = inspected.sourceFirst;
  while (current && !outlineRefEqual(current, inspected.managedContainer)) {
    const node = build(current);
    top.push(node);
    current = optionalOutlineReference(asLatin1(node.object.bytes), "Next", "PDF outline Next");
  }
  return top;
}

/** A bookmark's title as readers decode it (PDFDocEncoding, UTF-16, UTF-8);
 * null for a title held in another object. */
async function readSourceOutlineTitle(objectBytes: Uint8Array): Promise<string | null> {
  const source = asLatin1(objectBytes);
  const entry = oneTopLevelDictionaryEntry(source, "Title", "PDF outline Title", true);
  if (!entry) return "";
  const raw = source.slice(entry.valueStart, entry.valueEnd).trim();
  const { PDFHexString, PDFString } = await import("pdf-lib");
  const value = raw.startsWith("(") && raw.endsWith(")")
    ? PDFString.of(raw.slice(1, -1))
    : raw.startsWith("<") && raw.endsWith(">") && !raw.startsWith("<<")
      ? PDFHexString.of(raw.slice(1, -1))
      : null;
  if (!value) return null;
  return value.decodeText();
}

/** The file's own bookmarks (MRAI's subtree left out) as the AI names them:
 * each with its place in the outline (sibling indexes) and its title; a
 * bookmark whose title is held in another object is listed with an empty
 * title. What the API's bookmark command renames or deletes by title
 * (rename-source / delete-source), planned here instead of by the server. */
export async function readMraiNativeSourceOutlineListing(
  sourceBytes: Uint8Array,
): Promise<readonly Readonly<{ path: readonly number[]; title: string }>[]> {
  const bytes = Uint8Array.from(sourceBytes);
  const parsed = parsePdf(bytes, false);
  const catalog = parseIndirectObject(bytes, parsed, parsed.current.rootObjectNumber, parsed.current.rootGeneration);
  if (catalog.streamBytes !== null || !isCatalogObject(asLatin1(catalog.bytes))) {
    fail("unsupported-document", "PDF trailer Root is not a direct Catalog dictionary.");
  }
  const inspected = inspectOutlineTree(bytes, parsed, asLatin1(catalog.bytes));
  const listing: Readonly<{ path: readonly number[]; title: string }>[] = [];
  const walk = async (nodes: readonly SourceOutlineNode[], path: readonly number[]): Promise<void> => {
    for (const [index, node] of nodes.entries()) {
      const here = Object.freeze([...path, index]);
      listing.push(Object.freeze({ path: here, title: (await readSourceOutlineTitle(node.object.bytes)) ?? "" }));
      await walk(node.children, here);
    }
  };
  await walk(readSourceOutlineTree(bytes, parsed, inspected), []);
  return Object.freeze(listing);
}

/** The rows a bookmark reveals when open: its children, and the rows of each
 * child that is open itself (a positive /Count), as the inspectors count. */
function sourceVisibleWhenOpen(node: SourceOutlineNode): number {
  return node.children.reduce(
    (total, child) => total + 1 + ((child.count ?? 0) > 0 ? sourceVisibleWhenOpen(child) : 0),
    0,
  );
}

/** Rename or delete one of the file's own bookmarks: the object revisions it
 * takes and the top-level state the rest of the plan continues from. Mirror
 * of the save's `_source_outline_edit`. */
async function planSourceOutlineEdit(
  bytes: Uint8Array,
  parsed: ParsedPdf,
  inspected: InspectedOutlineTree,
  operation: "rename-source" | "delete-source",
  target: Readonly<{ path: readonly number[]; title: string | null; expectedTitle?: string }>,
): Promise<{
  readonly revisions: readonly IncrementalObjectRevision[];
  readonly sourceFirst: PdfNativeOutlineObjectRef | null;
  readonly sourceLast: PdfNativeOutlineObjectRef | null;
  readonly sourceVisibleCount: number;
  readonly sourceTarget: PdfNativeOutlineSourceTarget;
}> {
  const tree = readSourceOutlineTree(bytes, parsed, inspected);
  if (target.path.length === 0) fail("precondition-mismatch", "The bookmark is no longer in the outline.");
  let siblings = tree;
  const ancestors: SourceOutlineNode[] = [];
  for (const [depth, index] of target.path.entries()) {
    if (!Number.isInteger(index) || index < 0 || index >= siblings.length) {
      fail("precondition-mismatch", "The bookmark is no longer in the outline.");
    }
    if (depth === target.path.length - 1) break;
    ancestors.push(siblings[index]);
    siblings = siblings[index].children;
  }
  const index = target.path.at(-1)!;
  const node = siblings[index];
  if (target.expectedTitle !== undefined) {
    // The bookmark at that place still reads as the one the person chose
    // (as the outline pane shows it): never another one that moved there.
    const shown = await readSourceOutlineTitle(node.object.bytes);
    if (shown !== null && sanitizeOutlineTitle(shown) !== target.expectedTitle) {
      fail("precondition-mismatch", "The bookmark changed; reopen the outline.");
    }
  }
  const sourceTarget = Object.freeze({
    ...node.ref,
    expectedObjectSha256: await sha256PdfBytes(node.object.bytes),
    title: operation === "rename-source" ? target.title : null,
  });
  const updates = new Map<string, { node: SourceOutlineNode; changes: Record<string, string | null> }>();
  const update = (item: SourceOutlineNode, key: string, value: string | null) => {
    const identity = `${item.ref.objectNumber}:${item.ref.generation}`;
    const entry = updates.get(identity) ?? { node: item, changes: {} };
    entry.changes[key] = value;
    updates.set(identity, entry);
  };
  const refText = (ref: PdfNativeOutlineObjectRef | null) => ref ? `${ref.objectNumber} ${ref.generation} R` : null;
  if (operation === "rename-source") {
    if (!target.title) fail("invalid-request", "A bookmark rename names the new title.");
    update(node, "Title", outlineUtf16Hex(target.title));
  } else {
    siblings.splice(index, 1);
    const previous = index > 0 ? siblings[index - 1] : null;
    const following = index < siblings.length ? siblings[index] : null;
    if (previous) {
      // The last of the file's own top-level bookmarks leads into MRAI's subtree.
      update(previous, "Next", refText(following?.ref ?? (ancestors.length === 0 ? inspected.managedContainer : null)));
    }
    if (following) update(following, "Prev", refText(previous?.ref ?? null));
    const parent = ancestors.at(-1);
    if (parent) {
      if (siblings.length > 0) {
        update(parent, "First", refText(siblings[0].ref));
        update(parent, "Last", refText(siblings.at(-1)!.ref));
      } else {
        update(parent, "First", null);
        update(parent, "Last", null);
        update(parent, "Count", null);
        parent.count = null;
      }
    }
    for (const ancestor of [...ancestors].reverse()) {
      if (ancestor.children.length === 0 || ancestor.count === null || ancestor.count === 0) continue;
      const visible = sourceVisibleWhenOpen(ancestor);
      ancestor.count = ancestor.count > 0 ? visible : -visible;
      update(ancestor, "Count", String(ancestor.count));
    }
  }
  const revisions = [...updates.values()].map(({ node: item, changes }) =>
    Object.freeze({ ...item.ref, bytes: rewriteOutlineDictionaryObject(item.object, changes) }),
  );
  return Object.freeze({
    revisions,
    sourceFirst: tree[0]?.ref ?? null,
    sourceLast: tree.at(-1)?.ref ?? null,
    sourceVisibleCount: tree.reduce(
      (total, item) => total + 1 + ((item.count ?? 0) > 0 ? sourceVisibleWhenOpen(item) : 0),
      0,
    ),
    sourceTarget,
  });
}

/**
 * MRAI's own bookmarks as the file holds them (the subtree this planner and
 * the save's remap write): each node's id, title and destination on the
 * file's pages. The window reads its Bookmarks pane from here, so a reload,
 * a byte-published edit (whose record carries no tree) or a page edit's
 * remap all show what the file says. Refused (as the planner refuses) for a
 * subtree it cannot read exactly, including a destination on no page.
 */
export async function readMraiNativeManagedOutline(
  sourceBytes: Uint8Array,
): Promise<Readonly<{ pageCount: number; tree: readonly PdfNativeManagedOutlineNode[] }>> {
  const bytes = Uint8Array.from(sourceBytes);
  const parsed = parsePdf(bytes, false);
  const catalog = parseIndirectObject(bytes, parsed, parsed.current.rootObjectNumber, parsed.current.rootGeneration);
  if (catalog.streamBytes !== null || !isCatalogObject(asLatin1(catalog.bytes))) {
    fail("unsupported-document", "PDF trailer Root is not a direct Catalog dictionary.");
  }
  const inspected = inspectOutlineTree(bytes, parsed, asLatin1(catalog.bytes));
  const pages = readOutlinePageReferences(bytes, parsed, asLatin1(catalog.bytes));
  if (!inspected.managedContainer) return Object.freeze({ pageCount: pages.length, tree: Object.freeze([]) });
  const coordinate = (token: string): number | null => {
    if (token === "null") return null;
    const value = Number(token);
    if (!/^[+-]?(\d+\.?\d*|\.\d+)$/.test(token) || !Number.isFinite(value)) {
      fail("unsupported-document", "Managed outline destination is malformed.");
    }
    return value;
  };
  const build = async (ref: PdfNativeOutlineObjectRef, depth: number): Promise<PdfNativeManagedOutlineNode> => {
    if (depth > 12) fail("unsupported-document", "PDF outline depth exceeds 12.");
    const object = parseIndirectObject(bytes, parsed, ref.objectNumber, ref.generation);
    const source = asLatin1(object.bytes);
    const id = managedOutlineId(source);
    const title = await readSourceOutlineTitle(object.bytes);
    const dest = oneTopLevelDictionaryEntry(source, "Dest", "PDF outline Dest", true);
    const match = dest
      ? /^\[\s*(\d+)\s+(\d+)\s+R\s*\/XYZ\s+(\S+)\s+(\S+)\s+([^\s\]]+)\s*\]$/.exec(
        source.slice(dest.valueStart, dest.valueEnd).trim())
      : null;
    if (!id || title === null || !match) fail("unsupported-document", "Managed outline node is malformed.");
    const page = Object.freeze({ objectNumber: Number(match[1]), generation: Number(match[2]) });
    const index = pages.findIndex((candidate) => outlineRefEqual(candidate, page));
    if (index < 0) fail("unsupported-document", "Managed outline destination page is missing.");
    const children: PdfNativeManagedOutlineNode[] = [];
    let child = optionalOutlineReference(source, "First", "PDF outline First");
    while (child) {
      children.push(await build(child, depth + 1));
      const childSource = asLatin1(parseIndirectObject(bytes, parsed, child.objectNumber, child.generation).bytes);
      child = optionalOutlineReference(childSource, "Next", "PDF outline Next");
    }
    return Object.freeze({
      id: id!,
      title: title!,
      destination: Object.freeze({
        pageNumber: index + 1,
        page,
        left: coordinate(match[3]),
        top: coordinate(match[4]),
        zoom: coordinate(match[5]),
      }),
      children: Object.freeze(children),
    });
  };
  const containerSource = asLatin1(
    parseIndirectObject(bytes, parsed, inspected.managedContainer.objectNumber, inspected.managedContainer.generation).bytes,
  );
  const tree: PdfNativeManagedOutlineNode[] = [];
  let current = optionalOutlineReference(containerSource, "First", "PDF outline First");
  while (current) {
    tree.push(await build(current, 0));
    const currentSource = asLatin1(parseIndirectObject(bytes, parsed, current.objectNumber, current.generation).bytes);
    current = optionalOutlineReference(currentSource, "Next", "PDF outline Next");
  }
  return Object.freeze({ pageCount: pages.length, tree: Object.freeze(tree) });
}

export async function planMraiNativeOutlineMutation(input: {
  readonly sourceBytes: Uint8Array;
  readonly request: PdfNativeOutlineMutationRequest;
}): Promise<PdfNativePreparedOutlineMutation> {
  const sourceBytes = Uint8Array.from(input.sourceBytes);
  const sourceSha256 = await sha256PdfBytes(sourceBytes);
  if (sourceSha256 !== exactSha256(input.request.expectedSourceSha256, "source")) {
    fail("source-mismatch", "Managed outline source SHA-256 changed.");
  }
  // An xref-stream source needs no special handling here: every object this
  // planner reads is resolved through the parsed chain (object streams
  // included), and the revision it appends is one cross-reference stream
  // `incrementalXrefSection` builds from the stream dictionary, keeping
  // /Root, /Info and /ID. Requiring a classic source trailer only refused
  // pdfcpu, pdfTeX and LibreOffice documents the writers already support.
  const parsed = parsePdf(sourceBytes, false);
  const catalogRef = Object.freeze({
    objectNumber: parsed.current.rootObjectNumber,
    generation: parsed.current.rootGeneration,
  });
  const catalog = parseIndirectObject(
    sourceBytes,
    parsed,
    catalogRef.objectNumber,
    catalogRef.generation,
  );
  if (catalog.streamBytes !== null || !isCatalogObject(asLatin1(catalog.bytes))) {
    fail("unsupported-document", "PDF trailer Root is not a direct Catalog dictionary.");
  }
  const inspected = inspectOutlineTree(sourceBytes, parsed, asLatin1(catalog.bytes));
  if (
    input.request.expectedManagedIds.length !== inspected.managedIds.length ||
    input.request.expectedManagedIds.some((id, index) => id !== inspected.managedIds[index])
  ) {
    fail("precondition-mismatch", "Managed outline identity set changed.");
  }
  const pages = readOutlinePageReferences(sourceBytes, parsed, asLatin1(catalog.bytes));
  const pageMap = input.request.pageMap === undefined
    ? undefined
    : readOutlinePageMap(input.request.pageMap, pages.length);
  const managedTree = normalizeManagedOutlineInput(
    pageMap ? managedOutlineOnFilePages(input.request.managedTree, pageMap) : input.request.managedTree,
    pages,
  );
  const sourceOperation = input.request.operation === "rename-source" || input.request.operation === "delete-source"
    ? input.request.operation
    : null;
  if (sourceOperation && !input.request.sourceTarget) {
    fail("invalid-request", "A source bookmark edit names its bookmark.");
  }
  const sourceEdit = sourceOperation
    ? await planSourceOutlineEdit(sourceBytes, parsed, inspected, sourceOperation, input.request.sourceTarget!)
    : null;
  const outlineSource = sourceEdit ?? {
    sourceFirst: inspected.sourceFirst,
    sourceLast: inspected.sourceLast,
    sourceVisibleCount: inspected.sourceVisibleCount,
  };
  const latest = latestEntries(parsed);
  let nextObjectNumber = Math.max(parsed.current.size, Math.max(0, ...latest.keys()) + 1);
  const allocate = (): PdfNativeOutlineObjectRef =>
    Object.freeze({ objectNumber: nextObjectNumber++, generation: 0 });
  const rootRef = inspected.rootRef ?? allocate();
  const containerRef = managedTree.length > 0 ? inspected.managedContainer ?? allocate() : null;
  // The container keeps the title the file's already has (any language's);
  // a new one is named in its reader's language, English where the planner
  // runs with no page (the engine host): lib/pdf/managed-bookmarks-title.
  let containerTitle = pdfManagedBookmarksTitle(pageInterfaceLocale());
  if (containerRef && inspected.managedContainer) {
    const existingTitle = await readSourceOutlineTitle(
      parseIndirectObject(
        sourceBytes,
        parsed,
        inspected.managedContainer.objectNumber,
        inspected.managedContainer.generation,
      ).bytes,
    );
    if (isPdfManagedBookmarksTitle(existingTitle)) containerTitle = existingTitle!;
  }
  const refsById = new Map(inspected.managedRefsById);
  const assign = (nodes: readonly PdfNativeManagedOutlineNode[]) =>
    nodes.forEach((node) => {
      if (!refsById.has(node.id)) refsById.set(node.id, allocate());
      assign(node.children);
    });
  assign(managedTree);

  const revisions: IncrementalObjectRevision[] = [...(sourceEdit?.revisions ?? [])];
  const addRevision = (ref: PdfNativeOutlineObjectRef, bytes: Uint8Array) => {
    revisions.push(Object.freeze({ ...ref, bytes }));
  };
  if (!inspected.rootRef) {
    addRevision(
      catalogRef,
      rewriteOutlineDictionaryObject(catalog, {
        Outlines: `${rootRef.objectNumber} ${rootRef.generation} R`,
      }),
    );
  }
  const nowHasManaged = containerRef !== null;
  const previouslyHadManaged = inspected.managedContainer !== null;
  if (
    outlineSource.sourceLast &&
    nowHasManaged !== previouslyHadManaged
  ) {
    const sourceLastObject = parseIndirectObject(
      sourceBytes,
      parsed,
      outlineSource.sourceLast.objectNumber,
      outlineSource.sourceLast.generation,
    );
    addRevision(
      outlineSource.sourceLast,
      rewriteOutlineDictionaryObject(sourceLastObject, {
        Next: containerRef
          ? `${containerRef.objectNumber} ${containerRef.generation} R`
          : null,
      }),
    );
  }
  const managedDescendants = (node: PdfNativeManagedOutlineNode): number =>
    node.children.reduce(
      (total, child) => total + 1 + managedDescendants(child),
      0,
    );
  const managedCount = managedTree.reduce(
    (total, node) => total + 1 + managedDescendants(node),
    0,
  );
  const rootObject = inspected.rootObject;
  const rootChanges = {
    First: outlineSource.sourceFirst
      ? `${outlineSource.sourceFirst.objectNumber} ${outlineSource.sourceFirst.generation} R`
      : containerRef
        ? `${containerRef.objectNumber} ${containerRef.generation} R`
        : null,
    Last: containerRef
      ? `${containerRef.objectNumber} ${containerRef.generation} R`
      : outlineSource.sourceLast
        ? `${outlineSource.sourceLast.objectNumber} ${outlineSource.sourceLast.generation} R`
        : null,
    Count: String(
      outlineSource.sourceVisibleCount + (containerRef ? 1 + managedCount : 0),
    ),
  } as const;
  addRevision(
    rootRef,
    rootObject
      ? rewriteOutlineDictionaryObject(rootObject, rootChanges)
      : ascii(
          `${rootRef.objectNumber} ${rootRef.generation} obj\n<<\n/Type /Outlines\n${rootChanges.First ? `/First ${rootChanges.First}\n/Last ${rootChanges.Last}\n` : ""}/Count ${rootChanges.Count}\n>>\nendobj`,
        ),
  );

  const emitNodes = (
    nodes: readonly PdfNativeManagedOutlineNode[],
    parent: PdfNativeOutlineObjectRef,
  ) => {
    const refs = nodes.map((node) => refsById.get(node.id)!);
    nodes.forEach((node, index) => {
      const ref = refs[index];
      const childRefs = node.children.map((child) => refsById.get(child.id)!);
      addRevision(
        ref,
        serializeManagedOutlineNode(
          ref,
          parent,
          refs[index - 1] ?? null,
          refs[index + 1] ?? null,
          node,
          childRefs,
        ),
      );
      emitNodes(node.children, ref);
    });
  };
  if (containerRef) {
    const topRefs = managedTree.map((node) => refsById.get(node.id)!);
    const containerParts = [
      `${containerRef.objectNumber} ${containerRef.generation} obj`,
      "<<",
      `/Title ${outlineUtf16Hex(containerTitle)}`,
      `/Parent ${rootRef.objectNumber} ${rootRef.generation} R`,
    ];
    if (outlineSource.sourceLast) {
      containerParts.push(
        `/Prev ${outlineSource.sourceLast.objectNumber} ${outlineSource.sourceLast.generation} R`,
      );
    }
    if (topRefs.length > 0) {
      containerParts.push(`/First ${topRefs[0].objectNumber} ${topRefs[0].generation} R`);
      const last = topRefs.at(-1)!;
      containerParts.push(`/Last ${last.objectNumber} ${last.generation} R`);
      containerParts.push(`/Count ${managedCount}`);
    }
    containerParts.push("/MRAIManaged /Root", "/MRAISchema /V1");
    containerParts.push(">>", "endobj");
    addRevision(containerRef, ascii(containerParts.join("\n")));
    emitNodes(managedTree, containerRef);
  }
  const unique = new Map<string, IncrementalObjectRevision>();
  revisions.forEach((revision) => unique.set(`${revision.objectNumber}:${revision.generation}`, revision));
  const ordered = [...unique.values()].sort(
    (left, right) => left.objectNumber - right.objectNumber || left.generation - right.generation,
  );
  const incremental = makeMultiObjectIncrementalRevision(
    sourceBytes,
    parsed.current,
    ordered,
    { min: 1, max: 1_028, label: "Managed outline" },
  );
  const revisedObjects = await Promise.all(
    ordered.map(async (revision) =>
      Object.freeze({
        objectNumber: revision.objectNumber,
        generation: revision.generation,
        offset: incremental.offsets.get(`${revision.objectNumber}:${revision.generation}`)!,
        sha256: await sha256PdfBytes(revision.bytes),
      }),
    ),
  );
  return Object.freeze({
    evidence: Object.freeze({
      schemaVersion: MRAI_NATIVE_OUTLINE_SCHEMA_VERSION,
      operation: input.request.operation,
      expectedCatalog: Object.freeze({
        ...catalogRef,
        expectedObjectSha256: await sha256PdfBytes(catalog.bytes),
      }),
      expectedOutlineRoot: inspected.rootRef && inspected.rootObject
        ? Object.freeze({
            ...inspected.rootRef,
            expectedObjectSha256: await sha256PdfBytes(inspected.rootObject.bytes),
          })
        : null,
      managedTree,
      expectedManagedIds: Object.freeze([...inspected.managedIds]),
      ...(sourceEdit ? { sourceTarget: sourceEdit.sourceTarget } : {}),
      ...(pageMap ? { pageMap } : {}),
    }),
    sourceLength: sourceBytes.length,
    sourceXrefOffset: parsed.current.offset,
    incrementalXrefOffset: incremental.incrementalXrefOffset,
    candidateSha256: await sha256PdfBytes(incremental.candidateBytes),
    candidateBytes: incremental.candidateBytes,
    revisedObjects: Object.freeze(revisedObjects),
  });
}

// ---- WPS 分割页面: a page becomes several ----------------------------------
//
// The cut lines are drawn over the page as the reader shows it (after
// /Rotate). Every part is a page object holding a copy of the page's
// dictionary -- the same /Contents and /Resources by reference, nothing
// else copied -- with its own /MediaBox and /CropBox, inserted right after
// the page in the same /Kids. Split in place (WPS 删除原页面), the page object
// itself shows the first part, so bookmarks and links to the page land on
// it; keeping the original (保留原页面), the page stays as it was and every
// part is a copy after it. A merge takes the copies back out of the page
// tree and gives the page what the split took: the undo of a split. Every
// dictionary the revision writes is re-serialized one entry a line
// (`<<\n/Key value\n...\n>>`), the entries it keeps copied verbatim, so the
// server's `_prepare_page_split_candidate` writes the same bytes.

export type PdfNativePageSplitBoxKey =
  | "MediaBox"
  | "CropBox"
  | "TrimBox"
  | "BleedBox"
  | "ArtBox";

/** What a split writes on the pages it makes, so its merge can find them. */
export type PdfNativePageSplitMarkerKey =
  | "MRAISplitId"
  | "MRAISplitPart"
  | "MRAISplitOf";

const PAGE_SPLIT_BOX_KEYS: readonly PdfNativePageSplitBoxKey[] = [
  "MediaBox",
  "CropBox",
  "TrimBox",
  "BleedBox",
  "ArtBox",
];
const PAGE_SPLIT_MARKER_KEYS: readonly PdfNativePageSplitMarkerKey[] = [
  "MRAISplitId",
  "MRAISplitPart",
  "MRAISplitOf",
];
/** What a copy does not take from the page: the structure tree, article
 * threads, page actions, thumbnail and annotations stay the page's own. */
const PAGE_SPLIT_COPY_DROPPED_KEYS = [
  "StructParents",
  "B",
  "AA",
  "Thumb",
  "Annots",
] as const;
const PAGE_SPLIT_ID = /^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$/;
const PAGE_SPLIT_BOX_TEXT =
  /^\[(-?\d+(?:\.\d{1,4})?) (-?\d+(?:\.\d{1,4})?) (-?\d+(?:\.\d{1,4})?) (-?\d+(?:\.\d{1,4})?)\]$/;
/** A box entry a merge writes back: four numbers in one array, or a
 * reference to such an array -- never any other PDF syntax. */
const PAGE_SPLIT_RAW_BOX_VALUE =
  /^(?:\[[\x00\x09\x0a\x0c\x0d\x20]*[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[\x00\x09\x0a\x0c\x0d\x20]+[+-]?(?:\d+(?:\.\d*)?|\.\d+)){3}[\x00\x09\x0a\x0c\x0d\x20]*\]|\d+[\x00\x09\x0a\x0c\x0d\x20]+\d+[\x00\x09\x0a\x0c\x0d\x20]+R)$/;
/** A marker a merge writes back: as an earlier split wrote it. */
const PAGE_SPLIT_RAW_MARKER_VALUE: Readonly<Record<PdfNativePageSplitMarkerKey, RegExp>> = {
  MRAISplitId: /^\([A-Za-z0-9][A-Za-z0-9._-]{0,63}\)$/,
  MRAISplitPart: /^(?:[1-9]|1[0-6])$/,
  MRAISplitOf: /^[1-9]\d{0,9} \d{1,5} R$/,
};
const PAGE_SPLIT_BOX_NUMBER = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/;
const PAGE_SPLIT_MAX_REVISIONS = 1_028;
const PAGE_SPLIT_MIN_CELL = 0.01;
const PAGE_SPLIT_TOLERANCE = 0.001;

export interface PdfNativePageSplitGrid {
  /** Horizontal cut lines: fractions of the displayed height, from the top. */
  readonly rows: readonly number[];
  /** Vertical cut lines: fractions of the displayed width, from the left. */
  readonly columns: readonly number[];
}

/** A page's own box entries as raw PDF value text; null where it has none. */
export type PdfNativePageBoxEntries = Readonly<
  Record<PdfNativePageSplitBoxKey, string | null>
>;

/** A page's own split markers as raw PDF value text; null where it has none. */
export type PdfNativePageSplitMarkers = Readonly<
  Record<PdfNativePageSplitMarkerKey, string | null>
>;

export interface PdfNativePageSplitPageRequest {
  /** 1-based, in the source document. */
  readonly pageNumber: number;
  readonly grid: PdfNativePageSplitGrid;
}

export interface PdfNativePageSplitMergeGroup {
  readonly page: PdfNativeOutlineObjectRef;
  /** The copies the split wrote, in order. */
  readonly parts: readonly PdfNativeOutlineObjectRef[];
  readonly restoreBoxes: PdfNativePageBoxEntries;
  readonly restoreMarkers: PdfNativePageSplitMarkers;
}

export type PdfNativePageSplitRequest = Readonly<
  | {
      operation: "split";
      splitId: string;
      expectedSourceSha256: string;
      /** WPS 保留原页面: the page stays and every part is a copy after it. */
      keepOriginal: boolean;
      pages: readonly PdfNativePageSplitPageRequest[];
    }
  | {
      operation: "merge";
      splitId: string;
      expectedSourceSha256: string;
      keepOriginal: boolean;
      groups: readonly PdfNativePageSplitMergeGroup[];
    }
>;

export interface PdfNativePageSplitPart extends PdfNativeOutlineObjectRef {
  /** `[x0 y0 x1 y1]` in user space, as written. */
  readonly box: string;
}

export interface PdfNativePageSplitEvidenceGroup {
  readonly pageNumber: number;
  readonly page: PdfNativeOutlineObjectRef & {
    readonly expectedObjectSha256: string;
  };
  readonly parent: PdfNativeOutlineObjectRef;
  /** The page's own boxes before the split; for a merge, what it gets back. */
  readonly originalBoxes: PdfNativePageBoxEntries;
  /** The page's own markers (an earlier split's) before the split; for a
   * merge, what it gets back. */
  readonly originalMarkers: PdfNativePageSplitMarkers;
  /** Split: every part in reading order (in place, the page itself first);
   * merge: the copies taken away. */
  readonly parts: readonly PdfNativePageSplitPart[];
}

export interface PdfNativePageSplitEvidence {
  readonly schemaVersion: typeof MRAI_NATIVE_PAGE_SPLIT_SCHEMA_VERSION;
  readonly operation: "split" | "merge";
  readonly splitId: string;
  readonly keepOriginal: boolean;
  readonly groups: readonly PdfNativePageSplitEvidenceGroup[];
  readonly pageTreeNodes: readonly (PdfNativeOutlineObjectRef & {
    readonly expectedObjectSha256: string;
  })[];
}

export interface PdfNativePreparedPageSplit {
  readonly evidence: PdfNativePageSplitEvidence;
  readonly sourceLength: number;
  readonly sourceXrefOffset: number;
  readonly incrementalXrefOffset: number;
  readonly candidateSha256: string;
  readonly candidateBytes: Uint8Array;
  readonly revisedObjects: readonly Readonly<
    PdfNativeOutlineObjectRef & { offset: number; sha256: string }
  >[];
}

type PageSplitBox = readonly [number, number, number, number];

interface PageSplitTreeLeaf {
  readonly ref: PdfNativeOutlineObjectRef;
  readonly parent: PdfNativeOutlineObjectRef;
  readonly object: ParsedIndirectObject;
  readonly dictionary: string;
}

interface PageSplitTreeNode {
  readonly ref: PdfNativeOutlineObjectRef;
  readonly parent: PdfNativeOutlineObjectRef | null;
  readonly object: ParsedIndirectObject;
  readonly dictionary: string;
  readonly kids: readonly PdfNativeOutlineObjectRef[];
  readonly count: number;
}

interface PageSplitTree {
  readonly leaves: readonly PageSplitTreeLeaf[];
  readonly nodes: ReadonlyMap<string, PageSplitTreeNode>;
}

function pageSplitKey(ref: PdfNativeOutlineObjectRef): string {
  return `${ref.objectNumber}:${ref.generation}`;
}

function pageSplitRefText(ref: PdfNativeOutlineObjectRef): string {
  return `${ref.objectNumber} ${ref.generation} R`;
}

function refusePageSplit(
  code: PdfNativeTransactionErrorCode,
  message: string,
): never {
  throw new PdfNativeTransactionError(code, message);
}

/** The dictionary an indirect object holds, as the server reads it. */
function pageSplitDictionary(object: ParsedIndirectObject): string {
  if (object.streamBytes !== null) {
    refusePageSplit("unsupported-document", MRAI_NATIVE_PAGE_SPLIT_REFUSALS.pageTree);
  }
  const source = asLatin1(object.bytes);
  let start = 0;
  if (object.storage === "direct") {
    const header =
      /^\d+[\x00\x09\x0a\x0c\x0d\x20]+\d+[\x00\x09\x0a\x0c\x0d\x20]+obj\b/.exec(
        source,
      );
    if (!header) {
      refusePageSplit("unsupported-document", MRAI_NATIVE_PAGE_SPLIT_REFUSALS.pageTree);
    }
    start = skipWhitespace(source, header[0].length);
  } else {
    start = skipWhitespace(source, 0);
  }
  if (source.slice(start, start + 2) !== "<<") {
    refusePageSplit("unsupported-document", MRAI_NATIVE_PAGE_SPLIT_REFUSALS.pageTree);
  }
  return source.slice(start, dictionaryEnd(source, start));
}

/**
 * Re-serialize a dictionary one entry a line: entries in `updates` are
 * replaced (or dropped for null), every other entry is copied verbatim, and
 * update keys the dictionary did not have are appended in `updates` order.
 */
function pageSplitRewriteDictionary(
  dictionary: string,
  updates: readonly (readonly [string, string | null])[],
): string {
  const entries = topLevelDictionaryEntries(dictionary);
  const keys = new Set<string>();
  for (const entry of entries) {
    if (keys.has(entry.key)) {
      refusePageSplit("unsupported-document", MRAI_NATIVE_PAGE_SPLIT_REFUSALS.pageTree);
    }
    keys.add(entry.key);
  }
  const changes = new Map(updates);
  const lines: string[] = [];
  for (const entry of entries) {
    if (changes.has(entry.key)) {
      const value = changes.get(entry.key);
      if (value !== null && value !== undefined) {
        lines.push(`/${entry.key} ${value}`);
      }
      continue;
    }
    lines.push(dictionary.slice(entry.keyStart, entry.valueEnd));
  }
  for (const [key, value] of updates) {
    if (value !== null && !keys.has(key)) lines.push(`/${key} ${value}`);
  }
  return `<<\n${lines.join("\n")}\n>>`;
}

function pageSplitObjectBytes(
  ref: PdfNativeOutlineObjectRef,
  dictionary: string,
): Uint8Array {
  return latin1Bytes(
    `${ref.objectNumber} ${ref.generation} obj\n${dictionary}\nendobj`,
  );
}

/** One raw value of a page's own dictionary, or null where it has none. */
function pageSplitRawValue(dictionary: string, key: string): string | null {
  const entry = oneTopLevelDictionaryEntry(dictionary, key, `PDF page ${key}`, true);
  return entry ? dictionary.slice(entry.valueStart, entry.valueEnd) : null;
}

function pageSplitBoxNumbers(
  text: string,
  label: string,
): PageSplitBox {
  const trimmed = text.trim();
  if (!trimmed.startsWith("[") || !trimmed.endsWith("]")) {
    refusePageSplit("unsupported-document", `${label} is not one numeric array.`);
  }
  const tokens = trimmed
    .slice(1, -1)
    .split(/[\x00\x09\x0a\x0c\x0d\x20]+/)
    .filter((token) => token.length > 0);
  if (
    tokens.length !== 4 ||
    tokens.some((token) => !PAGE_SPLIT_BOX_NUMBER.test(token))
  ) {
    refusePageSplit("unsupported-document", `${label} is not four numbers.`);
  }
  const values = tokens.map(Number);
  if (
    values.some((value) => !Number.isFinite(value) || Math.abs(value) > 1_000_000) ||
    values[2] <= values[0] ||
    values[3] <= values[1]
  ) {
    refusePageSplit("unsupported-document", `${label} is empty or malformed.`);
  }
  return [values[0], values[1], values[2], values[3]];
}

/** A page box entry's numbers; an indirect array object is followed. */
function pageSplitBoxValue(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
  raw: string,
  label: string,
): PageSplitBox {
  const reference = /^(\d+)[\x00\x09\x0a\x0c\x0d\x20]+(\d+)[\x00\x09\x0a\x0c\x0d\x20]+R$/.exec(
    raw.trim(),
  );
  if (!reference) return pageSplitBoxNumbers(raw, label);
  const object = parseIndirectObject(
    sourceBytes,
    parsed,
    Number(reference[1]),
    Number(reference[2]),
  );
  if (object.streamBytes !== null) {
    refusePageSplit("unsupported-document", `${label} is not one numeric array.`);
  }
  let body = asLatin1(object.bytes);
  if (object.storage === "direct") {
    const header =
      /^[\x00\x09\x0a\x0c\x0d\x20]*\d+[\x00\x09\x0a\x0c\x0d\x20]+\d+[\x00\x09\x0a\x0c\x0d\x20]+obj\b/.exec(
        body,
      );
    const end = body.lastIndexOf("endobj");
    if (!header || end < header[0].length) {
      refusePageSplit("unsupported-document", `${label} is malformed.`);
    }
    body = body.slice(header[0].length, end);
  }
  return pageSplitBoxNumbers(body, label);
}

function readPageSplitTree(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
): PageSplitTree {
  const catalog = parseIndirectObject(
    sourceBytes,
    parsed,
    parsed.current.rootObjectNumber,
    parsed.current.rootGeneration,
  );
  if (catalog.streamBytes !== null || !isCatalogObject(asLatin1(catalog.bytes))) {
    refusePageSplit("unsupported-document", "PDF trailer Root is not a direct Catalog dictionary.");
  }
  const pagesRoot = oneTopLevelReference(
    pageSplitDictionary(catalog),
    "Pages",
    "PDF Catalog Pages reference",
  );
  if (!pagesRoot) {
    refusePageSplit("unsupported-document", "PDF Catalog Pages reference is missing.");
  }
  const seen = new Set<string>();
  const leaves: PageSplitTreeLeaf[] = [];
  const nodes = new Map<string, PageSplitTreeNode>();
  const visit = (
    ref: PdfNativeOutlineObjectRef,
    parent: PdfNativeOutlineObjectRef | null,
    depth: number,
  ): number => {
    if (depth > 32 || seen.size > 100_000) {
      refusePageSplit("unsupported-document", "PDF page tree is too deep or large.");
    }
    const key = pageSplitKey(ref);
    if (seen.has(key)) {
      refusePageSplit("unsupported-document", "PDF page tree is cyclic or duplicated.");
    }
    seen.add(key);
    const object = parseIndirectObject(
      sourceBytes,
      parsed,
      ref.objectNumber,
      ref.generation,
    );
    const dictionary = pageSplitDictionary(object);
    const declaredParent = oneTopLevelReference(
      dictionary,
      "Parent",
      "PDF page-tree Parent reference",
      true,
    );
    if (!outlineRefEqual(declaredParent, parent)) {
      refusePageSplit("unsupported-document", "PDF page-tree Parent identity is invalid.");
    }
    const type = exactObjectType(dictionary);
    if (type === "Page") {
      if (!parent) {
        refusePageSplit("unsupported-document", MRAI_NATIVE_PAGE_SPLIT_REFUSALS.pageTree);
      }
      leaves.push(
        Object.freeze({
          ref: Object.freeze({ ...ref }),
          parent: Object.freeze({ ...parent }),
          object,
          dictionary,
        }),
      );
      return 1;
    }
    if (type !== "Pages") {
      refusePageSplit("unsupported-document", "PDF page-tree node type is invalid.");
    }
    const kids = pageTreeKids(dictionary);
    const count = pageTreeCount(dictionary);
    const total = kids.reduce(
      (sum, kid) => sum + visit(kid, ref, depth + 1),
      0,
    );
    if (total !== count) {
      refusePageSplit("unsupported-document", "PDF page-tree Count is inconsistent.");
    }
    nodes.set(
      key,
      Object.freeze({
        ref: Object.freeze({ ...ref }),
        parent: parent ? Object.freeze({ ...parent }) : null,
        object,
        dictionary,
        kids: Object.freeze(kids.map((kid) => Object.freeze({ ...kid }))),
        count,
      }),
    );
    return total;
  };
  visit(pagesRoot, null, 0);
  return Object.freeze({ leaves: Object.freeze(leaves), nodes });
}

/** A page's inherited value: its own entry, else the nearest ancestor's. */
function pageSplitInherited(
  leaf: PageSplitTreeLeaf,
  tree: PageSplitTree,
  key: string,
): string | null {
  const own = pageSplitRawValue(leaf.dictionary, key);
  if (own !== null) return own;
  let cursor: PdfNativeOutlineObjectRef | null = leaf.parent;
  while (cursor) {
    const node = tree.nodes.get(pageSplitKey(cursor));
    if (!node) break;
    const value = pageSplitRawValue(node.dictionary, key);
    if (value !== null) return value;
    cursor = node.parent;
  }
  return null;
}

/** The part of the page a reader shows: CropBox within MediaBox. */
function pageSplitVisibleBox(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
  leaf: PageSplitTreeLeaf,
  tree: PageSplitTree,
): PageSplitBox {
  const mediaRaw = pageSplitInherited(leaf, tree, "MediaBox");
  if (mediaRaw === null) {
    refusePageSplit("unsupported-document", "PDF page MediaBox is missing.");
  }
  const media = pageSplitBoxValue(sourceBytes, parsed, mediaRaw, "PDF page MediaBox");
  const cropRaw = pageSplitInherited(leaf, tree, "CropBox");
  const crop = cropRaw === null
    ? media
    : pageSplitBoxValue(sourceBytes, parsed, cropRaw, "PDF page CropBox");
  const box: PageSplitBox = [
    Math.max(media[0], crop[0]),
    Math.max(media[1], crop[1]),
    Math.min(media[2], crop[2]),
    Math.min(media[3], crop[3]),
  ];
  if (box[2] - box[0] < 1 || box[3] - box[1] < 1) {
    refusePageSplit("unsupported-document", "PDF page shows no area to split.");
  }
  return box;
}

function pageSplitRotation(leaf: PageSplitTreeLeaf, tree: PageSplitTree): 0 | 90 | 180 | 270 {
  const raw = pageSplitInherited(leaf, tree, "Rotate");
  if (raw === null) return 0;
  const text = raw.trim();
  if (!/^[+-]?\d+$/.test(text)) {
    refusePageSplit("unsupported-document", "PDF page Rotate is not one integer.");
  }
  const value = Number(text);
  if (!Number.isSafeInteger(value) || value % 90 !== 0) {
    refusePageSplit("unsupported-document", "PDF page Rotate is not orthogonal.");
  }
  return (((value % 360) + 360) % 360) as 0 | 90 | 180 | 270;
}

/** A user-space coordinate as the record and the file write it. */
function pageSplitNumber(value: number): string {
  const rounded = Math.round(value * 10_000) / 10_000;
  return Object.is(rounded, -0) || rounded === 0 ? "0" : String(rounded);
}

function pageSplitBoxText(box: readonly [string, string, string, string]): string {
  return `[${box.join(" ")}]`;
}

function parsePageSplitBoxText(text: string): PageSplitBox {
  const match = PAGE_SPLIT_BOX_TEXT.exec(text);
  if (!match) {
    refusePageSplit("invalid-request", "Page split box is malformed.");
  }
  const box: PageSplitBox = [
    Number(match[1]),
    Number(match[2]),
    Number(match[3]),
    Number(match[4]),
  ];
  if (box[2] <= box[0] || box[3] <= box[1]) {
    refusePageSplit("invalid-request", "Page split box is empty.");
  }
  return box;
}

function pageSplitCuts(values: readonly number[], label: string): readonly number[] {
  if (
    !Array.isArray(values) ||
    values.length > PDF_NATIVE_PAGE_SPLIT_MAX_LINES ||
    values.some((value) => typeof value !== "number" || !Number.isFinite(value))
  ) {
    refusePageSplit("invalid-request", `${label} are invalid.`);
  }
  const cuts = [0, ...values, 1];
  for (let index = 1; index < cuts.length; index += 1) {
    if (cuts[index] - cuts[index - 1] < PAGE_SPLIT_MIN_CELL) {
      refusePageSplit("invalid-request", `${label} leave a part too narrow.`);
    }
  }
  return cuts;
}

/**
 * The parts of one page in reading order (rows top to bottom, each row left
 * to right, as displayed), in user space. A cut is written once and shared
 * by the parts on both sides, so the parts tile the page exactly.
 */
function pageSplitPartBoxes(
  box: PageSplitBox,
  rotation: 0 | 90 | 180 | 270,
  grid: PdfNativePageSplitGrid,
): readonly string[] {
  if (!grid || typeof grid !== "object") {
    refusePageSplit("invalid-request", "Page split lines are invalid.");
  }
  const uCuts = pageSplitCuts(grid.columns, "Page split columns");
  const vCuts = pageSplitCuts(grid.rows, "Page split rows");
  if ((uCuts.length - 1) * (vCuts.length - 1) < 2) {
    refusePageSplit("invalid-request", "A page split needs at least one cut.");
  }
  if ((uCuts.length - 1) * (vCuts.length - 1) > PDF_NATIVE_PAGE_SPLIT_MAX_PARTS) {
    refusePageSplit("invalid-request", MRAI_NATIVE_PAGE_SPLIT_REFUSALS.tooLarge);
  }
  const [x0, y0, x1, y1] = box;
  const width = x1 - x0;
  const height = y1 - y0;
  // Which user axis a display cut lands on, and where.
  const along = (fraction: number, axis: "u" | "v"): { axis: "x" | "y"; value: string } => {
    const edge = fraction === 0 || fraction === 1;
    switch (rotation) {
      case 90:
        return axis === "u"
          ? { axis: "y", value: pageSplitNumber(edge ? (fraction === 0 ? y0 : y1) : y0 + fraction * height) }
          : { axis: "x", value: pageSplitNumber(edge ? (fraction === 0 ? x0 : x1) : x0 + fraction * width) };
      case 180:
        return axis === "u"
          ? { axis: "x", value: pageSplitNumber(edge ? (fraction === 0 ? x1 : x0) : x1 - fraction * width) }
          : { axis: "y", value: pageSplitNumber(edge ? (fraction === 0 ? y0 : y1) : y0 + fraction * height) };
      case 270:
        return axis === "u"
          ? { axis: "y", value: pageSplitNumber(edge ? (fraction === 0 ? y1 : y0) : y1 - fraction * height) }
          : { axis: "x", value: pageSplitNumber(edge ? (fraction === 0 ? x1 : x0) : x1 - fraction * width) };
      default:
        return axis === "u"
          ? { axis: "x", value: pageSplitNumber(edge ? (fraction === 0 ? x0 : x1) : x0 + fraction * width) }
          : { axis: "y", value: pageSplitNumber(edge ? (fraction === 0 ? y1 : y0) : y1 - fraction * height) };
    }
  };
  const uLines = uCuts.map((fraction) => along(fraction, "u"));
  const vLines = vCuts.map((fraction) => along(fraction, "v"));
  const parts: string[] = [];
  for (let row = 0; row + 1 < vLines.length; row += 1) {
    for (let column = 0; column + 1 < uLines.length; column += 1) {
      const lines = [uLines[column], uLines[column + 1], vLines[row], vLines[row + 1]];
      const xs = lines.filter((line) => line.axis === "x").map((line) => line.value);
      const ys = lines.filter((line) => line.axis === "y").map((line) => line.value);
      const [minX, maxX] = Number(xs[0]) <= Number(xs[1]) ? [xs[0], xs[1]] : [xs[1], xs[0]];
      const [minY, maxY] = Number(ys[0]) <= Number(ys[1]) ? [ys[0], ys[1]] : [ys[1], ys[0]];
      parts.push(pageSplitBoxText([minX, minY, maxX, maxY]));
    }
  }
  return Object.freeze(parts);
}

/** The parts cover `box` exactly: each inside it, none overlapping, no gap. */
function assertPageSplitTiling(box: PageSplitBox, parts: readonly PageSplitBox[]): void {
  const area = (value: PageSplitBox) => (value[2] - value[0]) * (value[3] - value[1]);
  let covered = 0;
  for (const part of parts) {
    if (
      part[0] < box[0] - PAGE_SPLIT_TOLERANCE ||
      part[1] < box[1] - PAGE_SPLIT_TOLERANCE ||
      part[2] > box[2] + PAGE_SPLIT_TOLERANCE ||
      part[3] > box[3] + PAGE_SPLIT_TOLERANCE
    ) {
      refusePageSplit("invalid-request", "A page split part lies outside its page.");
    }
    covered += area(part);
  }
  for (let left = 0; left < parts.length; left += 1) {
    for (let right = left + 1; right < parts.length; right += 1) {
      const a = parts[left];
      const b = parts[right];
      const overlapX = Math.min(a[2], b[2]) - Math.max(a[0], b[0]);
      const overlapY = Math.min(a[3], b[3]) - Math.max(a[1], b[1]);
      if (overlapX > PAGE_SPLIT_TOLERANCE && overlapY > PAGE_SPLIT_TOLERANCE) {
        refusePageSplit("invalid-request", "Page split parts overlap.");
      }
    }
  }
  const whole = area(box);
  if (Math.abs(covered - whole) > Math.max(1, whole * 0.000_001)) {
    refusePageSplit("invalid-request", "Page split parts do not cover their page.");
  }
}

function pageSplitAnnotationsEmpty(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
  dictionary: string,
): boolean {
  const raw = pageSplitRawValue(dictionary, "Annots");
  if (raw === null) return true;
  let text = raw.trim();
  const reference = /^(\d+)[\x00\x09\x0a\x0c\x0d\x20]+(\d+)[\x00\x09\x0a\x0c\x0d\x20]+R$/.exec(text);
  if (reference) {
    const object = parseIndirectObject(
      sourceBytes,
      parsed,
      Number(reference[1]),
      Number(reference[2]),
    );
    if (object.streamBytes !== null) return false;
    text = asLatin1(object.bytes);
    if (object.storage === "direct") {
      const header =
        /^[\x00\x09\x0a\x0c\x0d\x20]*\d+[\x00\x09\x0a\x0c\x0d\x20]+\d+[\x00\x09\x0a\x0c\x0d\x20]+obj\b/.exec(text);
      const end = text.lastIndexOf("endobj");
      if (!header || end < header[0].length) return false;
      text = text.slice(header[0].length, end);
    }
    text = text.trim();
  }
  return /^\[[\x00\x09\x0a\x0c\x0d\x20]*\]$/.test(text);
}

/** Whether a request value is an array, without narrowing its declared
 * element type away (a request crosses the wire, so the shape is checked). */
function pageSplitIsList(value: unknown): boolean {
  return Array.isArray(value);
}

/** The page's own boxes, as a merge may write them back. */
function pageSplitOwnBoxes(dictionary: string): PdfNativePageBoxEntries {
  return Object.freeze(
    Object.fromEntries(
      PAGE_SPLIT_BOX_KEYS.map((key) => {
        const value = pageSplitRawValue(dictionary, key);
        if (value !== null && (value.length > 200 || !PAGE_SPLIT_RAW_BOX_VALUE.test(value))) {
          refusePageSplit("unsupported-document", MRAI_NATIVE_PAGE_SPLIT_REFUSALS.pageTree);
        }
        return [key, value];
      }),
    ) as Record<PdfNativePageSplitBoxKey, string | null>,
  );
}

/** The page's own split markers (an earlier split's), as a merge may write
 * them back. */
function pageSplitOwnMarkers(dictionary: string): PdfNativePageSplitMarkers {
  return Object.freeze(
    Object.fromEntries(
      PAGE_SPLIT_MARKER_KEYS.map((key) => {
        const value = pageSplitRawValue(dictionary, key);
        if (value !== null && !PAGE_SPLIT_RAW_MARKER_VALUE[key].test(value)) {
          refusePageSplit("unsupported-document", MRAI_NATIVE_PAGE_SPLIT_REFUSALS.pageTree);
        }
        return [key, value];
      }),
    ) as Record<PdfNativePageSplitMarkerKey, string | null>,
  );
}

/** What a merge was asked to write back, held to the shapes a split reads. */
function pageSplitRestoreEntries(
  group: PdfNativePageSplitMergeGroup,
): Readonly<{ boxes: PdfNativePageBoxEntries; markers: PdfNativePageSplitMarkers }> {
  const boxes = Object.fromEntries(
    PAGE_SPLIT_BOX_KEYS.map((key) => {
      const value = group.restoreBoxes?.[key] ?? null;
      if (
        value !== null &&
        (typeof value !== "string" || value.length > 200 || !PAGE_SPLIT_RAW_BOX_VALUE.test(value))
      ) {
        refusePageSplit("invalid-request", `PDF page ${key} cannot be written back.`);
      }
      return [key, value];
    }),
  ) as Record<PdfNativePageSplitBoxKey, string | null>;
  const markers = Object.fromEntries(
    PAGE_SPLIT_MARKER_KEYS.map((key) => {
      const value = group.restoreMarkers?.[key] ?? null;
      if (value !== null && (typeof value !== "string" || !PAGE_SPLIT_RAW_MARKER_VALUE[key].test(value))) {
        refusePageSplit("invalid-request", `PDF page ${key} cannot be written back.`);
      }
      return [key, value];
    }),
  ) as Record<PdfNativePageSplitMarkerKey, string | null>;
  return Object.freeze({ boxes: Object.freeze(boxes), markers: Object.freeze(markers) });
}

/** Plan a split of pages into parts, or the merge that undoes one. */
export async function planMraiNativePageSplit(input: {
  readonly sourceBytes: Uint8Array;
  readonly request: PdfNativePageSplitRequest;
}): Promise<PdfNativePreparedPageSplit> {
  const sourceBytes = Uint8Array.from(input.sourceBytes);
  const request = input.request;
  if (
    !request ||
    (request.operation !== "split" && request.operation !== "merge") ||
    typeof request.splitId !== "string" ||
    !PAGE_SPLIT_ID.test(request.splitId) ||
    typeof request.keepOriginal !== "boolean"
  ) {
    refusePageSplit("invalid-request", "Page split request is invalid.");
  }
  const keepOriginal = request.keepOriginal;
  const splitIdText = `(${request.splitId})`;
  const sourceSha256 = await sha256PdfBytes(sourceBytes);
  if (sourceSha256 !== exactSha256(request.expectedSourceSha256, "source")) {
    refusePageSplit("source-mismatch", "Page split source SHA-256 changed.");
  }
  const parsed = parsePdf(sourceBytes, false);
  const tree = readPageSplitTree(sourceBytes, parsed);
  const leafIndex = new Map(
    tree.leaves.map((leaf, index) => [pageSplitKey(leaf.ref), index] as const),
  );
  let highestObjectNumber = 0;
  for (const objectNumber of latestEntries(parsed).keys()) {
    if (objectNumber > highestObjectNumber) highestObjectNumber = objectNumber;
  }
  let nextObjectNumber = Math.max(parsed.current.size, highestObjectNumber + 1);
  const groups: PdfNativePageSplitEvidenceGroup[] = [];
  const revisions = new Map<string, IncrementalObjectRevision>();
  // Pages nodes the revision rewrites: the change in their page count, and
  // for a direct parent the kids it now lists.
  const countDelta = new Map<string, number>();
  const kidsEdits = new Map<string, Map<string, readonly PdfNativeOutlineObjectRef[]>>();
  const removedKids = new Map<string, Set<string>>();
  const addDelta = (parent: PdfNativeOutlineObjectRef, delta: number) => {
    let cursor: PdfNativeOutlineObjectRef | null = parent;
    while (cursor) {
      const key = pageSplitKey(cursor);
      countDelta.set(key, (countDelta.get(key) ?? 0) + delta);
      cursor = tree.nodes.get(key)?.parent ?? null;
    }
  };
  let newObjects = 0;

  if (request.operation === "split") {
    const pages = request.pages;
    if (
      !pageSplitIsList(pages) ||
      pages.length < 1 ||
      pages.length > PDF_NATIVE_PAGE_SPLIT_MAX_PAGES ||
      pages.some(
        (page, index) =>
          !page ||
          !Number.isSafeInteger(page.pageNumber) ||
          page.pageNumber < 1 ||
          page.pageNumber > tree.leaves.length ||
          (index > 0 && page.pageNumber <= pages[index - 1].pageNumber),
      )
    ) {
      refusePageSplit("invalid-request", "Page split pages are invalid.");
    }
    for (const { pageNumber, grid } of pages) {
      const leaf = tree.leaves[pageNumber - 1];
      if (!pageSplitAnnotationsEmpty(sourceBytes, parsed, leaf.dictionary)) {
        refusePageSplit("unsupported-document", MRAI_NATIVE_PAGE_SPLIT_REFUSALS.annotations);
      }
      const visible = pageSplitVisibleBox(sourceBytes, parsed, leaf, tree);
      const boxes = pageSplitPartBoxes(visible, pageSplitRotation(leaf, tree), grid);
      assertPageSplitTiling(visible, boxes.map(parsePageSplitBoxText));
      const originalBoxes = pageSplitOwnBoxes(leaf.dictionary);
      const originalMarkers = pageSplitOwnMarkers(leaf.dictionary);
      const parts: PdfNativePageSplitPart[] = [];
      boxes.forEach((box, index) => {
        const inPlace = !keepOriginal && index === 0;
        const ref = inPlace
          ? leaf.ref
          : Object.freeze({ objectNumber: nextObjectNumber++, generation: 0 });
        if (!inPlace) newObjects += 1;
        parts.push(Object.freeze({ ...ref, box }));
        const ownBoxUpdates = PAGE_SPLIT_BOX_KEYS.filter(
          (key) => key === "MediaBox" || key === "CropBox" || originalBoxes[key] !== null,
        ).map((key) => [key, box] as const);
        const updates: (readonly [string, string | null])[] = inPlace
          ? [
              ...ownBoxUpdates,
              ["MRAISplitId", splitIdText],
              ["MRAISplitPart", "1"],
              ["MRAISplitOf", null],
            ]
          : [
              ...ownBoxUpdates,
              ...PAGE_SPLIT_COPY_DROPPED_KEYS.map((key) => [key, null] as const),
              ["MRAISplitId", splitIdText],
              ["MRAISplitPart", String(index + 1)],
              ["MRAISplitOf", pageSplitRefText(leaf.ref)],
            ];
        revisions.set(
          pageSplitKey(ref),
          Object.freeze({
            ...ref,
            bytes: pageSplitObjectBytes(ref, pageSplitRewriteDictionary(leaf.dictionary, updates)),
          }),
        );
      });
      const copies = keepOriginal ? parts : parts.slice(1);
      const parentKey = pageSplitKey(leaf.parent);
      if (!kidsEdits.has(parentKey)) kidsEdits.set(parentKey, new Map());
      kidsEdits.get(parentKey)!.set(
        pageSplitKey(leaf.ref),
        copies.map((part) => Object.freeze({
          objectNumber: part.objectNumber,
          generation: part.generation,
        })),
      );
      addDelta(leaf.parent, copies.length);
      groups.push(
        Object.freeze({
          pageNumber,
          page: Object.freeze({
            ...leaf.ref,
            expectedObjectSha256: await sha256PdfBytes(leaf.object.bytes),
          }),
          parent: leaf.parent,
          originalBoxes,
          originalMarkers,
          parts: Object.freeze(parts),
        }),
      );
    }
  } else {
    const requestedGroups = request.groups;
    if (
      !pageSplitIsList(requestedGroups) ||
      requestedGroups.length < 1 ||
      requestedGroups.length > PDF_NATIVE_PAGE_SPLIT_MAX_PAGES
    ) {
      refusePageSplit("invalid-request", "Page merge groups are invalid.");
    }
    const seenParts = new Set<string>();
    const seenPages = new Set<string>();
    const firstCopy = keepOriginal ? 1 : 2;
    for (const group of requestedGroups) {
      const index = group ? leafIndex.get(pageSplitKey(group.page)) : undefined;
      if (
        index === undefined ||
        seenPages.has(pageSplitKey(group.page)) ||
        !pageSplitIsList(group.parts) ||
        group.parts.length < (keepOriginal ? 2 : 1) ||
        group.parts.length > PDF_NATIVE_PAGE_SPLIT_MAX_PARTS - (keepOriginal ? 0 : 1)
      ) {
        refusePageSplit("precondition-mismatch", MRAI_NATIVE_PAGE_SPLIT_REFUSALS.notSplit);
      }
      seenPages.add(pageSplitKey(group.page));
      const leaf = tree.leaves[index];
      const parent = tree.nodes.get(pageSplitKey(leaf.parent))!;
      const position = parent.kids.findIndex((kid) => outlineRefEqual(kid, leaf.ref));
      if (
        !keepOriginal &&
        (pageSplitRawValue(leaf.dictionary, "MRAISplitId") !== splitIdText ||
          pageSplitRawValue(leaf.dictionary, "MRAISplitPart") !== "1" ||
          pageSplitRawValue(leaf.dictionary, "MRAISplitOf") !== null)
      ) {
        refusePageSplit("precondition-mismatch", MRAI_NATIVE_PAGE_SPLIT_REFUSALS.notSplit);
      }
      const removed: PdfNativePageSplitPart[] = [];
      const partBoxes: PageSplitBox[] = keepOriginal
        ? []
        : [parsePageSplitBoxText(pageSplitRawValue(leaf.dictionary, "CropBox") ?? "")];
      group.parts.forEach((partRef, partIndex) => {
        const kid = parent.kids[position + 1 + partIndex];
        const partLeafIndex = leafIndex.get(pageSplitKey(partRef));
        if (
          !kid ||
          !outlineRefEqual(kid, partRef) ||
          partLeafIndex === undefined ||
          seenParts.has(pageSplitKey(partRef)) ||
          seenPages.has(pageSplitKey(partRef))
        ) {
          refusePageSplit("precondition-mismatch", MRAI_NATIVE_PAGE_SPLIT_REFUSALS.notSplit);
        }
        seenParts.add(pageSplitKey(partRef));
        const part = tree.leaves[partLeafIndex];
        if (
          pageSplitRawValue(part.dictionary, "MRAISplitId") !== splitIdText ||
          pageSplitRawValue(part.dictionary, "MRAISplitPart") !== String(partIndex + firstCopy) ||
          pageSplitRawValue(part.dictionary, "MRAISplitOf") !== pageSplitRefText(leaf.ref)
        ) {
          refusePageSplit("precondition-mismatch", MRAI_NATIVE_PAGE_SPLIT_REFUSALS.notSplit);
        }
        // Undoing the split takes the copy away: refuse rather than lose
        // the comments put on it since.
        if (!pageSplitAnnotationsEmpty(sourceBytes, parsed, part.dictionary)) {
          refusePageSplit("unsupported-document", MRAI_NATIVE_PAGE_SPLIT_REFUSALS.partAnnotations);
        }
        const box = pageSplitRawValue(part.dictionary, "CropBox") ?? "";
        partBoxes.push(parsePageSplitBoxText(box));
        removed.push(Object.freeze({ ...part.ref, box }));
      });
      // The page after the last copy is no part of this split: the merge
      // takes back exactly the pages the split wrote.
      const after = parent.kids[position + 1 + group.parts.length];
      if (after) {
        const afterLeaf = tree.leaves[leafIndex.get(pageSplitKey(after)) ?? -1];
        if (
          afterLeaf &&
          pageSplitRawValue(afterLeaf.dictionary, "MRAISplitOf") === pageSplitRefText(leaf.ref) &&
          pageSplitRawValue(afterLeaf.dictionary, "MRAISplitId") === splitIdText
        ) {
          refusePageSplit("precondition-mismatch", MRAI_NATIVE_PAGE_SPLIT_REFUSALS.notSplit);
        }
      }
      const restore = pageSplitRestoreEntries(group);
      if (keepOriginal) {
        // The split never touched the page: the record names what it has,
        // and the copies showed it all.
        const ownBoxes = pageSplitOwnBoxes(leaf.dictionary);
        const ownMarkers = pageSplitOwnMarkers(leaf.dictionary);
        if (
          PAGE_SPLIT_BOX_KEYS.some((key) => ownBoxes[key] !== restore.boxes[key]) ||
          PAGE_SPLIT_MARKER_KEYS.some((key) => ownMarkers[key] !== restore.markers[key])
        ) {
          refusePageSplit("precondition-mismatch", MRAI_NATIVE_PAGE_SPLIT_REFUSALS.notSplit);
        }
        assertPageSplitTiling(pageSplitVisibleBox(sourceBytes, parsed, leaf, tree), partBoxes);
      } else {
        for (const key of PAGE_SPLIT_BOX_KEYS) {
          const value = restore.boxes[key];
          if (value !== null) pageSplitBoxValue(sourceBytes, parsed, value, `PDF page ${key}`);
        }
        const restoredDictionary = pageSplitRewriteDictionary(leaf.dictionary, [
          ...PAGE_SPLIT_BOX_KEYS.map((key) => [key, restore.boxes[key]] as const),
          ...PAGE_SPLIT_MARKER_KEYS.map((key) => [key, restore.markers[key]] as const),
        ]);
        // The boxes written back show what the parts showed together.
        const restoredLeaf: PageSplitTreeLeaf = { ...leaf, dictionary: restoredDictionary };
        assertPageSplitTiling(
          pageSplitVisibleBox(sourceBytes, parsed, restoredLeaf, tree),
          partBoxes,
        );
        revisions.set(
          pageSplitKey(leaf.ref),
          Object.freeze({
            ...leaf.ref,
            bytes: pageSplitObjectBytes(leaf.ref, restoredDictionary),
          }),
        );
      }
      const parentKey = pageSplitKey(leaf.parent);
      if (!removedKids.has(parentKey)) removedKids.set(parentKey, new Set());
      for (const part of removed) removedKids.get(parentKey)!.add(pageSplitKey(part));
      addDelta(leaf.parent, -removed.length);
      groups.push(
        Object.freeze({
          pageNumber: index + 1,
          page: Object.freeze({
            ...leaf.ref,
            expectedObjectSha256: await sha256PdfBytes(leaf.object.bytes),
          }),
          parent: leaf.parent,
          originalBoxes: restore.boxes,
          originalMarkers: restore.markers,
          parts: Object.freeze(removed),
        }),
      );
    }
  }
  if (newObjects > PDF_NATIVE_PAGE_SPLIT_MAX_PAGES * 4) {
    refusePageSplit("invalid-request", MRAI_NATIVE_PAGE_SPLIT_REFUSALS.tooLarge);
  }

  const pageTreeNodes: (PdfNativeOutlineObjectRef & { expectedObjectSha256: string })[] = [];
  for (const [key, delta] of countDelta) {
    const node = tree.nodes.get(key);
    if (!node) {
      refusePageSplit("unsupported-document", MRAI_NATIVE_PAGE_SPLIT_REFUSALS.pageTree);
    }
    const updates: (readonly [string, string | null])[] = [];
    const inserted = kidsEdits.get(key);
    const removed = removedKids.get(key);
    if (inserted || removed) {
      const kids = node.kids.flatMap((kid) =>
        removed?.has(pageSplitKey(kid))
          ? []
          : [kid, ...(inserted?.get(pageSplitKey(kid)) ?? [])],
      );
      if (kids.length < 1 || kids.length > 4_096) {
        refusePageSplit("unsupported-document", MRAI_NATIVE_PAGE_SPLIT_REFUSALS.tooLarge);
      }
      updates.push(["Kids", `[${kids.map(pageSplitRefText).join(" ")}]`]);
    }
    const count = node.count + delta;
    if (count < 1 || count > 100_000) {
      refusePageSplit("unsupported-document", MRAI_NATIVE_PAGE_SPLIT_REFUSALS.tooLarge);
    }
    updates.push(["Count", String(count)]);
    revisions.set(
      key,
      Object.freeze({
        ...node.ref,
        bytes: pageSplitObjectBytes(node.ref, pageSplitRewriteDictionary(node.dictionary, updates)),
      }),
    );
    pageTreeNodes.push(
      Object.freeze({
        ...node.ref,
        expectedObjectSha256: await sha256PdfBytes(node.object.bytes),
      }),
    );
  }
  pageTreeNodes.sort(
    (left, right) =>
      left.objectNumber - right.objectNumber || left.generation - right.generation,
  );
  const ordered = [...revisions.values()].sort(
    (left, right) =>
      left.objectNumber - right.objectNumber || left.generation - right.generation,
  );
  const incremental = makeMultiObjectIncrementalRevision(
    sourceBytes,
    parsed.current,
    ordered,
    { min: 1, max: PAGE_SPLIT_MAX_REVISIONS, label: "Page split" },
  );
  // The written page tree is the one planned: the pages in their order, the
  // parts right after their page (or gone again), every count consistent.
  const reopened = readPageSplitTree(
    incremental.candidateBytes,
    parsePdf(incremental.candidateBytes, true),
  );
  const removedKeys = new Set(
    request.operation === "merge"
      ? groups.flatMap((item) => item.parts.map(pageSplitKey))
      : [],
  );
  const expected = tree.leaves.flatMap((leaf) => {
    if (request.operation === "merge") {
      return removedKeys.has(pageSplitKey(leaf.ref)) ? [] : [pageSplitKey(leaf.ref)];
    }
    const group = groups.find((item) => outlineRefEqual(item.page, leaf.ref));
    if (!group) return [pageSplitKey(leaf.ref)];
    const partKeys = group.parts.map((part) => pageSplitKey(part));
    return keepOriginal ? [pageSplitKey(leaf.ref), ...partKeys] : partKeys;
  });
  const actual = reopened.leaves.map((leaf) => pageSplitKey(leaf.ref));
  if (
    actual.length !== expected.length ||
    actual.some((key, index) => key !== expected[index])
  ) {
    refusePageSplit("verification-failed", "Reopened PDF page tree does not match the split.");
  }
  const revisedObjects = await Promise.all(
    ordered.map(async (revision) =>
      Object.freeze({
        objectNumber: revision.objectNumber,
        generation: revision.generation,
        offset: incremental.offsets.get(pageSplitKey(revision))!,
        sha256: await sha256PdfBytes(revision.bytes),
      }),
    ),
  );
  return Object.freeze({
    evidence: Object.freeze({
      schemaVersion: MRAI_NATIVE_PAGE_SPLIT_SCHEMA_VERSION,
      operation: request.operation,
      splitId: request.splitId,
      keepOriginal,
      groups: Object.freeze(groups),
      pageTreeNodes: Object.freeze(pageTreeNodes),
    }),
    sourceLength: sourceBytes.length,
    sourceXrefOffset: parsed.current.offset,
    incrementalXrefOffset: incremental.incrementalXrefOffset,
    candidateSha256: await sha256PdfBytes(incremental.candidateBytes),
    candidateBytes: incremental.candidateBytes,
    revisedObjects: Object.freeze(revisedObjects),
  });
}

// ---------------------------------------------------------------------------
// Page background (WPS 背景): a colour or a picture painted beneath what each
// page draws, written into the PDF as its own native revision.

/** The page's background fill, as the planner writes it. */
export type PdfNativePageBackgroundFillRequest =
  | Readonly<{ kind: "color"; color: readonly [number, number, number] }>
  | Readonly<{
      kind: "image";
      mimeType: "image/png" | "image/jpeg";
      imageBase64: string;
      /** The picture's opacity, 0.05..1 (1 or absent: opaque). */
      opacity?: number;
      /** `stretch` fills the page; `fit` keeps the picture's ratio, centred. */
      fit?: "stretch" | "fit";
      /** A colour filling the page under the picture. */
      color?: readonly [number, number, number];
    }>;

export type PdfNativePageBackgroundRequest =
  | Readonly<{
      operation: "apply";
      expectedSourceSha256: string;
      /** The pages that take the background, in document order. */
      pageNumbers: readonly number[];
      fill: PdfNativePageBackgroundFillRequest;
    }>
  | Readonly<{
      /** Undo or redo: each page goes back to the background it names
       * (an earlier one still in the file, or none). */
      operation: "restore";
      expectedSourceSha256: string;
      pages: readonly Readonly<{
        pageNumber: number;
        background: PdfNativeOutlineObjectRef | null;
      }>[];
    }>;

export type PdfNativePageBackgroundFillEvidence =
  | Readonly<{ kind: "color"; color: readonly [number, number, number] }>
  | Readonly<{
      kind: "image";
      image: PdfNativeEmbeddedImageReplacementData;
      reference: Readonly<{ objectNumber: number; generation: 0 }>;
      softMask: Readonly<{ objectNumber: number; generation: 0 }> | null;
      opacity: number | null;
      extGState: Readonly<{ objectNumber: number; generation: 0 }> | null;
      fit: "stretch" | "fit";
      color: readonly [number, number, number] | null;
    }>;

export interface PdfNativePageBackgroundPageEvidence {
  readonly pageNumber: number;
  readonly page: PdfNativeOutlineObjectRef & { readonly expectedObjectSha256: string };
  /** The part of the page a reader shows (CropBox within MediaBox). */
  readonly box: readonly [number, number, number, number];
  /** The background stream the page draws first after the write; null for none. */
  readonly background: PdfNativeOutlineObjectRef | null;
  /** Whether the write adds `background` (apply) or names one already in the file (restore). */
  readonly backgroundIsNew: boolean;
  /** The background stream the page drew first before the write, which it no longer draws. */
  readonly replaced: PdfNativeOutlineObjectRef | null;
  /** The page's new Resources (a picture background names its picture there). */
  readonly resources: PdfNativeOutlineObjectRef | null;
  readonly xObjectName: string | null;
  readonly extGStateName: string | null;
}

export interface PdfNativePageBackgroundEvidence {
  readonly schemaVersion: typeof MRAI_NATIVE_PAGE_BACKGROUND_SCHEMA_VERSION;
  readonly operation: "apply" | "restore";
  readonly fill: PdfNativePageBackgroundFillEvidence | null;
  readonly pages: readonly PdfNativePageBackgroundPageEvidence[];
}

export interface PdfNativePreparedPageBackground {
  readonly evidence: PdfNativePageBackgroundEvidence;
  readonly sourceLength: number;
  readonly sourceXrefOffset: number;
  readonly incrementalXrefOffset: number;
  readonly candidateSha256: string;
  readonly candidateBytes: Uint8Array;
  readonly revisedObjects: readonly Readonly<
    PdfNativeOutlineObjectRef & { offset: number; sha256: string }
  >[];
  

  readonly pdfaIdentificationRemoval?: PdfNativePreparedPdfaIdentificationRemoval;
  /** The catalog stating the version the revision needs (PDF 0020). */
  readonly catalogVersionRaise?: PdfNativePreparedCatalogVersionRaise;
}

const PAGE_BACKGROUND_MARKER = "MRAIBackground";

function refusePageBackground(code: PdfNativeTransactionErrorCode, message: string): never {
  throw new PdfNativeTransactionError(code, message);
}

function pageBackgroundRound(value: number): number {
  const rounded = Math.round(value * 1e6) / 1e6;
  return Object.is(rounded, -0) ? 0 : rounded;
}

/** Whether a stream is a background this writer made (its dictionary
 * carries the marker). */
function pageBackgroundStreamOf(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
  reference: PdfNativeOutlineObjectRef,
): boolean {
  let object: ParsedIndirectObject;
  try {
    object = parseIndirectObject(sourceBytes, parsed, reference.objectNumber, reference.generation);
  } catch (error) {
    if (error instanceof PdfNativeTransactionError) return false;
    throw error;
  }
  if (object.storage !== "direct" || object.streamBytes === null) return false;
  const dictionary = topLevelDictionarySource(asLatin1(object.bytes), "PDF page background stream");
  const marker = oneTopLevelDictionaryEntry(dictionary, PAGE_BACKGROUND_MARKER, "PDF page background marker", true);
  return marker !== null && dictionary.slice(marker.valueStart, marker.valueEnd).trim() === "true";
}

/** The block a background stream draws (mirror of the server's
 * `_page_background_block`). */
function pageBackgroundBlock(
  fill: PdfNativePageBackgroundFillEvidence,
  box: readonly [number, number, number, number],
  xObjectName: string | null,
  extGStateName: string | null,
): string {
  const [x0, y0, x1, y1] = box;
  const width = formatNativePdfNumber(x1 - x0);
  const height = formatNativePdfNumber(y1 - y0);
  const x = formatNativePdfNumber(x0);
  const y = formatNativePdfNumber(y0);
  const rect = (color: readonly [number, number, number]) => {
    const [r, g, b] = color.map(formatNativePdfNumber);
    return [`${r} ${g} ${b} rg`, `${x} ${y} ${width} ${height} re`, "f"];
  };
  // Each block ends at a line end: the page's next content stream is read
  // as its continuation, and a token must not run into the next stream's.
  if (fill.kind === "color") return `${["q", ...rect(fill.color), "Q"].join("\n")}\n`;
  // `fit`: the picture's own ratio, as large as the page holds, centred.
  let placed = [x1 - x0, y1 - y0, x0, y0];
  if (fill.fit === "fit") {
    const scale = Math.min((x1 - x0) / fill.image.width, (y1 - y0) / fill.image.height);
    const w = fill.image.width * scale;
    const h = fill.image.height * scale;
    placed = [w, h, x0 + ((x1 - x0) - w) / 2, y0 + ((y1 - y0) - h) / 2];
  }
  const [pw, ph, px, py] = placed.map(formatNativePdfNumber);
  return [
    "q",
    ...(extGStateName ? [`/${extGStateName} gs`] : []),
    ...(fill.color ? rect(fill.color) : []),
    `${pw} 0 0 ${ph} ${px} ${py} cm`,
    `/${xObjectName} Do`,
    "Q",
  ].join("\n") + "\n";
}

function pageBackgroundStreamObject(reference: PdfNativeOutlineObjectRef, block: string): Uint8Array {
  return ascii(
    `${reference.objectNumber} ${reference.generation} obj\n<< /Length ${block.length} /${PAGE_BACKGROUND_MARKER} true >>\nstream\n${block}\nendstream\nendobj`,
  );
}

function pageBackgroundFreeName(
  graph: ResolvedPageFontGraph,
  stem: string,
): string {
  for (let index = 0; index < 1000; index += 1) {
    const name = index === 0 ? stem : `${stem}${index}`;
    if (!graph.names.has(name) && !graph.entries.some((entry) => entry.resourceName === name)) return name;
  }
  refusePageBackground("unsupported-document", "PDF page resource names are exhausted.");
}

/**
 * WPS 背景 written into the PDF: one appended revision that makes each
 * target page draw a background stream before everything it drew (a colour
 * filling its visible box, or a picture stretched over it at an opacity),
 * the page's own content streams kept byte for byte. A page that already
 * begins with a background this writer made draws the new one instead.
 * `restore` puts each page back to the background it names (Undo/Redo).
 * Mirror of the server's `_prepare_page_background_candidate`.
 */
export async function planMraiNativePageBackground(input: {
  readonly sourceBytes: Uint8Array;
  readonly request: PdfNativePageBackgroundRequest;
}): Promise<PdfNativePreparedPageBackground> {
  const sourceBytes = Uint8Array.from(input.sourceBytes);
  const request = input.request;
  if (!request || (request.operation !== "apply" && request.operation !== "restore")) {
    refusePageBackground("invalid-request", "Page background request is invalid.");
  }
  if ((await sha256PdfBytes(sourceBytes)) !== exactSha256(request.expectedSourceSha256, "source")) {
    refusePageBackground("source-mismatch", "Page background source SHA-256 changed.");
  }
  const parsed = parsePdf(sourceBytes, false);
  const tree = readPageSplitTree(sourceBytes, parsed);
  const graphs = await inspectMraiNativeDocumentPageGraphs(sourceBytes);
  const wanted: readonly Readonly<{ pageNumber: number; background: PdfNativeOutlineObjectRef | null }>[] =
    request.operation === "apply"
      ? request.pageNumbers.map((pageNumber) => ({ pageNumber, background: null }))
      : request.pages;
  if (
    !Array.isArray(wanted) ||
    wanted.length < 1 ||
    wanted.length > PDF_NATIVE_PAGE_BACKGROUND_MAX_PAGES ||
    wanted.some(
      (item, index) =>
        !item ||
        !Number.isSafeInteger(item.pageNumber) ||
        item.pageNumber < 1 ||
        item.pageNumber > tree.leaves.length ||
        item.pageNumber > graphs.length ||
        (index > 0 && item.pageNumber <= wanted[index - 1].pageNumber),
    )
  ) {
    refusePageBackground("invalid-request", "Page background pages are invalid.");
  }
  let highestObjectNumber = 0;
  for (const objectNumber of latestEntries(parsed).keys()) {
    if (objectNumber > highestObjectNumber) highestObjectNumber = objectNumber;
  }
  let nextObjectNumber = Math.max(parsed.current.size, highestObjectNumber + 1);
  const allocate = () => Object.freeze({ objectNumber: nextObjectNumber++, generation: 0 as const });
  const revisions: IncrementalObjectRevision[] = [];

  // The picture and its opacity state are written once and named by every page.
  let fill: PdfNativePageBackgroundFillEvidence | null = null;
  if (request.operation === "apply") {
    const requested = request.fill;
    if (requested?.kind === "color") {
      const color = requested.color;
      if (
        !Array.isArray(color) ||
        color.length !== 3 ||
        color.some((value) => typeof value !== "number" || !Number.isFinite(value) || value < 0 || value > 1)
      ) {
        refusePageBackground("invalid-request", "Page background colour is invalid.");
      }
      fill = Object.freeze({
        kind: "color" as const,
        color: Object.freeze(color.map(pageBackgroundRound)) as unknown as readonly [number, number, number],
      });
    } else if (requested?.kind === "image") {
      const opacity = requested.opacity;
      if (
        opacity !== undefined &&
        (typeof opacity !== "number" || !Number.isFinite(opacity) || opacity < 0.05 || opacity > 1)
      ) {
        refusePageBackground("invalid-request", "Page background opacity is invalid.");
      }
      const image = await inspectMraiNativeEmbeddedImageReplacement(
        base64ToBytes(requested.imageBase64),
        requested.mimeType,
      );
      const reference = allocate();
      const softMask = image.softMask ? allocate() : null;
      const translucent = opacity !== undefined && pageBackgroundRound(opacity) < 1;
      const under = requested.color;
      if (
        (requested.fit !== undefined && requested.fit !== "stretch" && requested.fit !== "fit") ||
        (under !== undefined &&
          (!Array.isArray(under) || under.length !== 3 ||
            under.some((value) => typeof value !== "number" || !Number.isFinite(value) || value < 0 || value > 1)))
      ) {
        refusePageBackground("invalid-request", "Page background picture is invalid.");
      }
      const extGState = translucent ? allocate() : null;
      fill = Object.freeze({
        kind: "image" as const,
        image,
        reference,
        softMask,
        opacity: translucent ? pageBackgroundRound(opacity!) : null,
        extGState,
        fit: requested.fit ?? "stretch",
        color: under
          ? (Object.freeze(under.map(pageBackgroundRound)) as unknown as readonly [number, number, number])
          : null,
      });
      revisions.push(
        ...contentAppendAdditionObjects({ category: "XObject", resourceName: "MRAIBg", reference, softMask, image }),
      );
      if (extGState) {
        revisions.push(
          ...contentAppendAdditionObjects({
            category: "ExtGState",
            resourceName: "MRAIBgA",
            reference: extGState,
            fillOpacity: pageBackgroundRound(opacity!),
            strokeOpacity: null,
          }),
        );
      }
    } else {
      refusePageBackground("invalid-request", "Page background fill is invalid.");
    }
  }

  const pages: PdfNativePageBackgroundPageEvidence[] = [];
  for (const item of wanted) {
    const leaf = tree.leaves[item.pageNumber - 1];
    const graph = graphs[item.pageNumber - 1];
    if (!sameNativeObjectRef(graph.page, leaf.ref)) {
      refusePageBackground("unsupported-document", "PDF page tree is ambiguous.");
    }
    const box = pageSplitVisibleBox(sourceBytes, parsed, leaf, tree);
    const contents = graph.contentStreams;
    const replaced =
      contents.length > 0 && pageBackgroundStreamOf(sourceBytes, parsed, contents[0])
        ? Object.freeze({ objectNumber: contents[0].objectNumber, generation: contents[0].generation })
        : null;
    // Taking backgrounds away from pages that have none leaves those pages
    // out of the write (Remove on the whole document).
    if (!fill && !item.background && !replaced) continue;
    const rest = replaced ? contents.slice(1) : contents;
    let background: PdfNativeOutlineObjectRef | null = null;
    let resources: PdfNativeOutlineObjectRef | null = null;
    let xObjectName: string | null = null;
    let extGStateName: string | null = null;
    const updates: (readonly [string, string | null])[] = [];
    if (fill) {
      background = allocate();
      if (fill.kind === "image") {
        const xobjects = await resolvePageResourceCategoryGraph(sourceBytes, parsed, graph, "XObject");
        xObjectName = pageBackgroundFreeName(xobjects, "MRAIBg");
        let resourcesDictionary = xobjects.resourcesDictionary;
        resourcesDictionary = rewriteDictionaryEntry(
          resourcesDictionary,
          "XObject",
          rewriteDictionaryEntry(
            xobjects.categoryDictionary,
            xObjectName,
            `${fill.reference.objectNumber} 0 R`,
            "PDF XObject resource name",
          ),
          "PDF XObject category",
        );
        if (fill.extGState) {
          const states = await resolvePageResourceCategoryGraph(sourceBytes, parsed, graph, "ExtGState");
          extGStateName = pageBackgroundFreeName(states, "MRAIBgA");
          resourcesDictionary = rewriteDictionaryEntry(
            resourcesDictionary,
            "ExtGState",
            rewriteDictionaryEntry(
              states.categoryDictionary,
              extGStateName,
              `${fill.extGState.objectNumber} 0 R`,
              "PDF ExtGState resource name",
            ),
            "PDF ExtGState category",
          );
        }
        resources = allocate();
        revisions.push(
          Object.freeze({
            ...resources,
            bytes: latin1Bytes(`${resources.objectNumber} 0 obj\n${resourcesDictionary}\nendobj`),
          }),
        );
        updates.push(["Resources", `${resources.objectNumber} 0 R`]);
      }
      revisions.push(
        Object.freeze({
          ...background,
          bytes: pageBackgroundStreamObject(background, pageBackgroundBlock(fill, box, xObjectName, extGStateName)),
        }),
      );
    } else if (item.background) {
      if (!pageBackgroundStreamOf(sourceBytes, parsed, item.background)) {
        refusePageBackground("precondition-mismatch", "The background to restore is not one this writer made.");
      }
      background = Object.freeze({ objectNumber: item.background.objectNumber, generation: item.background.generation });
    }
    if (!replaced && !background) {
      refusePageBackground("precondition-mismatch", "The page has no background to take away.");
    }
    if (background && rest.some((reference) => sameNativeObjectRef(reference, background!))) {
      refusePageBackground("precondition-mismatch", "The background is already drawn by the page.");
    }
    const drawn = [...(background ? [background] : []), ...rest];
    updates.unshift(["Contents", `[${drawn.map((reference) => `${reference.objectNumber} ${reference.generation} R`).join(" ")}]`]);
    revisions.push(
      Object.freeze({
        ...leaf.ref,
        bytes: pageSplitObjectBytes(leaf.ref, pageSplitRewriteDictionary(leaf.dictionary, updates)),
      }),
    );
    pages.push(
      Object.freeze({
        pageNumber: item.pageNumber,
        page: Object.freeze({ ...leaf.ref, expectedObjectSha256: await sha256PdfBytes(leaf.object.bytes) }),
        box: Object.freeze([...box]) as unknown as readonly [number, number, number, number],
        background,
        backgroundIsNew: Boolean(fill),
        replaced,
        resources,
        xObjectName,
        extGStateName,
      }),
    );
  }
  if (pages.length === 0) {
    refusePageBackground("precondition-mismatch", "The pages have no background to take away.");
  }
  const ordered = [...revisions].sort(
    (left, right) => left.objectNumber - right.objectNumber || left.generation - right.generation,
  );
  const incremental = makeMultiObjectIncrementalRevision(sourceBytes, parsed.current, ordered, {
    min: 1,
    max: PDF_NATIVE_PAGE_BACKGROUND_MAX_PAGES * 3 + 3,
    label: "Page background",
  });
  // Over a source that claims PDF/A the catalog's packet is written last in
  // the same revision without the identification, as every native write
  
  const pdfaIdentification = await appendPdfaIdentificationRemoval(sourceBytes, parsed, incremental);
  // A background below 100% (or a picture with an alpha channel) in a file
  // declaring PDF 1.3: the catalog states 1.4, last in the same revision
  // (user audit PDF 0020).
  const catalogVersion = await appendCatalogVersionRaise(sourceBytes, parsed, pdfaIdentification ?? incremental);
  const candidateBytes = catalogVersion?.candidateBytes ?? pdfaIdentification?.candidateBytes ?? incremental.candidateBytes;
  // Every page draws, first, exactly the background planned.
  const reopened = await inspectMraiNativeDocumentPageGraphs(candidateBytes, {
    validateContentStreams: false,
    tolerateUnusablePages: true,
  });
  for (const page of pages) {
    const first = reopened[page.pageNumber - 1]?.contentStreams[0] ?? null;
    if (page.background ? !first || !sameNativeObjectRef(first, page.background) : first !== null && page.replaced !== null && sameNativeObjectRef(first, page.replaced)) {
      refusePageBackground("verification-failed", "Reopened PDF page does not draw the planned background.");
    }
  }
  const revisedObjects = await Promise.all(
    ordered.map(async (revision) =>
      Object.freeze({
        objectNumber: revision.objectNumber,
        generation: revision.generation,
        offset: incremental.offsets.get(`${revision.objectNumber}:${revision.generation}`)!,
        sha256: await sha256PdfBytes(revision.bytes),
      }),
    ),
  );
  for (const beside of [pdfaIdentification?.removal, catalogVersion?.raise]) {
    if (!beside) continue;
    const { objectNumber, generation, offset, sha256 } = beside;
    revisedObjects.push(Object.freeze({ objectNumber, generation, offset, sha256 }));
    revisedObjects.sort((left, right) => left.objectNumber - right.objectNumber || left.generation - right.generation);
  }
  return Object.freeze({
    evidence: Object.freeze({
      schemaVersion: MRAI_NATIVE_PAGE_BACKGROUND_SCHEMA_VERSION,
      operation: request.operation,
      fill,
      pages: Object.freeze(pages),
    }),
    sourceLength: sourceBytes.length,
    sourceXrefOffset: parsed.current.offset,
    incrementalXrefOffset:
      catalogVersion?.incrementalXrefOffset ?? pdfaIdentification?.incrementalXrefOffset ?? incremental.incrementalXrefOffset,
    candidateSha256: await sha256PdfBytes(candidateBytes),
    candidateBytes,
    revisedObjects: Object.freeze(revisedObjects),
    ...(pdfaIdentification ? { pdfaIdentificationRemoval: pdfaIdentification.removal } : {}),
    ...(catalogVersion ? { catalogVersionRaise: catalogVersion.raise } : {}),
  });
}

function normalizeFontExtensionPackage(
  value: PdfNativeFontExtensionPackage,
): PdfNativeFontExtensionPackage {
  if (
    value.schemaVersion !== MRAI_NATIVE_FONT_EXTENSION_PACKAGE_SCHEMA_VERSION ||
    (value.programFace === undefined
      // The bundled CJK fallback, or a packaged complex-script face a
      // shaped run is set in (buildPdfNativeShapedFontPackage).
      ? value.fontSha256 !== PDF_UNICODE_FALLBACK_FONT_SHA256 && !isPdfComplexScriptFontSha256(value.fontSha256)
      : !/^[0-9a-f]{64}$/u.test(value.fontSha256) ||
        value.fontSha256 === PDF_UNICODE_FALLBACK_FONT_SHA256 ||
        typeof value.programFace.family !== "string" ||
        !value.programFace.family ||
        value.programFace.family.length > 120 ||
        !Number.isSafeInteger(value.programFace.weight) ||
        value.programFace.weight < 1 ||
        value.programFace.weight > 1_000 ||
        typeof value.programFace.italic !== "boolean" ||
        (value.programFace.page !== undefined && value.programFace.page !== true)) ||
    !PDF_RESOURCE_NAME.test(value.resourceName) ||
    !Number.isSafeInteger(value.font.objectNumber) ||
    value.font.objectNumber < 1 ||
    value.font.generation !== 0 ||
    !/^[0-9A-F]+$/.test(value.encodedHex) ||
    value.encodedHex.length === 0 ||
    value.encodedHex.length > 16_000 ||
    value.encodedHex.length % 2 !== 0 ||
    !Number.isFinite(value.widthAt1000) ||
    value.widthAt1000 <= 0 ||
    !Number.isSafeInteger(value.unitsPerEm) ||
    value.unitsPerEm <= 0 ||
    !Number.isFinite(value.ascent) ||
    !Number.isFinite(value.descent) ||
    !Number.isSafeInteger(value.objectByteLength) ||
    value.objectByteLength <= 0 ||
    !Array.isArray(value.objects) ||
    value.objects.length < 4 ||
    value.objects.length > 16
  ) {
    fail("invalid-request", "Native font extension package is invalid.");
  }
  let byteLength = 0;
  const objects = value.objects.map((object, index) => {
    if (
      !Number.isSafeInteger(object.objectNumber) ||
      object.objectNumber !== value.font.objectNumber + index ||
      object.generation !== 0 ||
      !isUint8Array(object.bytes) ||
      object.bytes.length < 24 ||
      object.bytes.length > 16 * 1024 * 1024
    ) {
      fail("invalid-request", "Native font extension object is invalid.");
    }
    const bytes = Uint8Array.from(object.bytes);
    const source = asLatin1(bytes);
    if (
      !source.startsWith(`${object.objectNumber} 0 obj`) ||
      !/endobj\s*$/.test(source)
    ) {
      fail("invalid-request", "Native font extension object is malformed.");
    }
    byteLength += bytes.length;
    return Object.freeze({
      objectNumber: object.objectNumber,
      generation: 0 as const,
      sha256: exactSha256(object.sha256, "font extension object"),
      bytes,
    });
  });
  if (byteLength !== value.objectByteLength || byteLength > 16 * 1024 * 1024) {
    fail("invalid-request", "Native font extension byte length is invalid.");
  }
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_FONT_EXTENSION_PACKAGE_SCHEMA_VERSION,
    fontSha256: value.fontSha256,
    ...(value.programFace
      ? {
          programFace: Object.freeze({
            family: value.programFace.family,
            weight: value.programFace.weight,
            italic: value.programFace.italic,
            // A font from the reader's computer the page keeps.
            ...(value.programFace.page === true ? { page: true as const } : {}),
          }),
        }
      : {}),
    resourceName: value.resourceName,
    font: Object.freeze({
      objectNumber: value.font.objectNumber,
      generation: 0 as const,
    }),
    encodedHex: value.encodedHex,
    widthAt1000: value.widthAt1000,
    unitsPerEm: value.unitsPerEm,
    ascent: value.ascent,
    descent: value.descent,
    objectByteLength: byteLength,
    objects: Object.freeze(objects),
  });
}

function normalizeFontAliasPackage(
  value: PdfNativeFontAliasPackage,
): PdfNativeFontAliasPackage {
  if (
    value.schemaVersion !== MRAI_NATIVE_FONT_ALIAS_PACKAGE_SCHEMA_VERSION ||
    value.fontSha256 !== PDF_UNICODE_FALLBACK_FONT_SHA256 ||
    !PDF_RESOURCE_NAME.test(value.resourceName) ||
    !Number.isSafeInteger(value.font.objectNumber) ||
    value.font.objectNumber < 1 ||
    value.font.generation !== 0 ||
    !/^(?:[0-9A-F]{4})+$/.test(value.encodedHex) ||
    !Number.isFinite(value.widthAt1000) ||
    value.widthAt1000 <= 0 ||
    !Array.isArray(value.objects) ||
    value.objects.length !== 2
  ) {
    fail("invalid-request", "Native font alias package is invalid.");
  }
  for (const reference of [value.sourceFont, value.sourceToUnicode]) {
    assertObjectReference(reference.objectNumber, reference.generation);
    exactSha256(reference.expectedObjectSha256, "font alias source object");
  }
  let byteLength = 0;
  const objects = value.objects.map((object, index) => {
    if (
      object.objectNumber !== value.font.objectNumber + index ||
      object.generation !== 0 ||
      !isUint8Array(object.bytes) ||
      object.bytes.length < 24 ||
      object.bytes.length > 256 * 1024
    ) {
      fail("invalid-request", "Native font alias object is invalid.");
    }
    const bytes = Uint8Array.from(object.bytes);
    const source = asLatin1(bytes);
    if (
      !source.startsWith(`${object.objectNumber} 0 obj`) ||
      !/endobj\s*$/.test(source)
    ) {
      fail("invalid-request", "Native font alias object is malformed.");
    }
    byteLength += bytes.length;
    return Object.freeze({
      objectNumber: object.objectNumber,
      generation: 0 as const,
      sha256: exactSha256(object.sha256, "font alias object"),
      bytes,
    });
  });
  if (byteLength !== value.objectByteLength || byteLength > 256 * 1024) {
    fail("invalid-request", "Native font alias byte length is invalid.");
  }
  return Object.freeze({
    ...value,
    sourceFont: Object.freeze({ ...value.sourceFont }),
    sourceToUnicode: Object.freeze({ ...value.sourceToUnicode }),
    objects: Object.freeze(objects),
  });
}

function normalizePageFontPackage(
  value: PdfNativeFontExtensionPackage | PdfNativeFontAliasPackage,
) {
  return value.schemaVersion === MRAI_NATIVE_FONT_ALIAS_PACKAGE_SCHEMA_VERSION
    ? normalizeFontAliasPackage(value)
    : normalizeFontExtensionPackage(value);
}

/** The most fonts one write installs beside its first: a line set in a
 * served CJK face draws from one package per slice it needs. */
const MAX_ADDITIONAL_FONT_PACKAGES = 31;

/** The packages installed beside the first: served-face subsets only, each
 * numbered on from the end of the one before, every resource name its own. */
function normalizeAdditionalFontPackages(
  first: PdfNativeFontExtensionPackage | PdfNativeFontAliasPackage,
  values: readonly PdfNativeFontExtensionPackage[] | undefined,
): readonly PdfNativeFontExtensionPackage[] | undefined {
  if (values === undefined) return undefined;
  if (
    !Array.isArray(values) ||
    values.length < 1 ||
    values.length > MAX_ADDITIONAL_FONT_PACKAGES ||
    first.schemaVersion !== MRAI_NATIVE_FONT_EXTENSION_PACKAGE_SCHEMA_VERSION ||
    first.programFace === undefined
  ) {
    fail("invalid-request", "Additional font packages are invalid.");
  }
  const names = new Set([first.resourceName]);
  let next = first.objects.at(-1)!.objectNumber + 1;
  return Object.freeze(values.map((value) => {
    const normalized = normalizeFontExtensionPackage(value);
    if (
      normalized.programFace === undefined ||
      names.has(normalized.resourceName) ||
      normalized.font.objectNumber !== next
    ) {
      fail("invalid-request", "Additional font packages are invalid.");
    }
    names.add(normalized.resourceName);
    next = normalized.objects.at(-1)!.objectNumber + 1;
    return normalized;
  }));
}

const CONTENT_APPEND_CATEGORIES: readonly PdfNativeContentAppendCategory[] =
  Object.freeze(["Font", "XObject", "ExtGState"]);
const CONTENT_APPEND_CONTROL_CHARACTERS = /[\u0000-\u001f\u007f]/u;

function normalizeContentAppendCategoryState(
  value: PdfNativePageExtGStateInspection["category"],
  label: string,
): PdfNativePageExtGStateInspection["category"] {
  const expectedDictionarySha256 = exactSha256(
    value.expectedDictionarySha256,
    `${label} dictionary`,
  );
  if (value.kind === "missing" || value.kind === "direct") {
    return Object.freeze({ kind: value.kind, expectedDictionarySha256 });
  }
  if (value.kind !== "indirect") {
    fail("invalid-request", `${label} category kind is invalid.`);
  }
  assertObjectReference(value.target.objectNumber, value.target.generation);
  return Object.freeze({
    kind: "indirect" as const,
    expectedDictionarySha256,
    target: Object.freeze({
      objectNumber: value.target.objectNumber,
      generation: value.target.generation,
      expectedObjectSha256: exactSha256(
        value.target.expectedObjectSha256,
        `${label} category object`,
      ),
    }),
  });
}

function exactContentAppendOpacity(
  value: number | null,
  label: string,
): number | null {
  if (value === null) return null;
  if (
    typeof value !== "number" ||
    !Number.isFinite(value) ||
    value < 0 ||
    value > 1
  ) {
    fail("invalid-request", `${label} opacity is invalid.`);
  }
  return value;
}

function exactContentAppendNewRef(
  value: Readonly<{ objectNumber: number; generation: number }>,
  label: string,
): Readonly<{ objectNumber: number; generation: 0 }> {
  assertObjectReference(value.objectNumber, value.generation);
  if (value.generation !== 0) {
    fail("invalid-request", `${label} generation must be zero.`);
  }
  return Object.freeze({
    objectNumber: value.objectNumber,
    generation: 0 as const,
  });
}

function normalizeContentAppendPlan(
  plan: PdfNativeContentAppendPlan | undefined,
): PdfNativeContentAppendPlan | undefined {
  if (!plan) return undefined;
  if (plan.schemaVersion !== MRAI_NATIVE_CONTENT_APPEND_SCHEMA_VERSION) {
    fail("invalid-request", "Content append schema is unsupported.");
  }
  if (plan.direction !== "append" && plan.direction !== "remove") {
    fail("invalid-request", "Content append direction is invalid.");
  }
  if (
    plan.appendKind !== "text" &&
    plan.appendKind !== "image" &&
    plan.appendKind !== "path"
  ) {
    fail("invalid-request", "Content append kind is invalid.");
  }
  const pageNumber = exactPositiveInteger(
    plan.pageNumber,
    "content append page number",
  );
  if (pageNumber > 100_000) {
    fail("invalid-request", "Content append page number is too large.");
  }
  assertObjectReference(plan.page.objectNumber, plan.page.generation);
  assertObjectReference(
    plan.resourceOwner.objectNumber,
    plan.resourceOwner.generation,
  );
  if (
    plan.resourceOwner.dictionaryKind !== "nested-resources" &&
    plan.resourceOwner.dictionaryKind !== "resource-dictionary"
  ) {
    fail("invalid-request", "Content append Resources kind is invalid.");
  }
  if (
    !Number.isSafeInteger(plan.restoredDepth) ||
    plan.restoredDepth < 0 ||
    plan.restoredDepth > 64
  ) {
    fail("invalid-request", "Content append restored depth is invalid.");
  }
  const categories: Partial<
    Record<
      PdfNativeContentAppendCategory,
      PdfNativePageExtGStateInspection["category"]
    >
  > = {};
  for (const key of Object.keys(plan.categories)) {
    const category = key as PdfNativeContentAppendCategory;
    if (!CONTENT_APPEND_CATEGORIES.includes(category)) {
      fail("invalid-request", "Content append resource category is invalid.");
    }
    const state = plan.categories[category];
    if (!state) continue;
    categories[category] = normalizeContentAppendCategoryState(
      state,
      `Content append ${category}`,
    );
  }
  const claimed = new Set<string>();
  const claim = (category: PdfNativeContentAppendCategory, name: string) => {
    if (typeof name !== "string" || !PDF_RESOURCE_NAME.test(name)) {
      fail("invalid-request", "Content append resource name is invalid.");
    }
    const key = `${category}/${name}`;
    if (claimed.has(key)) {
      fail("invalid-request", "Content append resource name is duplicated.");
    }
    claimed.add(key);
  };
  const additions = plan.additions.map(
    (addition): PdfNativeContentAppendAddition => {
      if (!CONTENT_APPEND_CATEGORIES.includes(addition.category)) {
        fail("invalid-request", "Content append addition category is invalid.");
      }
      claim(addition.category, addition.resourceName);
      if (!categories[addition.category]) {
        fail("invalid-request", "Content append category state is missing.");
      }
      if (addition.category === "Font") {
        const fontPackage = normalizeFontExtensionPackage(addition.fontPackage);
        if (
          fontPackage.resourceName !== addition.resourceName ||
          typeof addition.text !== "string" ||
          !addition.text ||
          addition.text.length > 2_000 ||
          CONTENT_APPEND_CONTROL_CHARACTERS.test(addition.text)
        ) {
          fail("invalid-request", "Content append font addition is invalid.");
        }
        return Object.freeze({
          category: "Font" as const,
          resourceName: addition.resourceName,
          text: addition.text,
          fontPackage,
        });
      }
      if (addition.category === "XObject") {
        const normalized = normalizedEmbeddedImageReplacementData(
          addition.image,
        );
        const reference = exactContentAppendNewRef(
          addition.reference,
          "Content append image",
        );
        const softMask = addition.softMask
          ? exactContentAppendNewRef(
              addition.softMask,
              "Content append soft mask",
            )
          : null;
        if (
          Boolean(softMask) !== Boolean(normalized.softMaskBytes) ||
          (softMask && softMask.objectNumber !== reference.objectNumber + 1)
        ) {
          fail("invalid-request", "Content append image allocation is invalid.");
        }
        return Object.freeze({
          category: "XObject" as const,
          resourceName: addition.resourceName,
          reference,
          softMask,
          image: normalized.evidence,
        });
      }
      const fillOpacity = exactContentAppendOpacity(
        addition.fillOpacity,
        "Content append fill",
      );
      const strokeOpacity = exactContentAppendOpacity(
        addition.strokeOpacity,
        "Content append stroke",
      );
      if (fillOpacity === null && strokeOpacity === null) {
        fail("invalid-request", "Content append ExtGState addition is empty.");
      }
      return Object.freeze({
        category: "ExtGState" as const,
        resourceName: addition.resourceName,
        reference: exactContentAppendNewRef(
          addition.reference,
          "Content append ExtGState",
        ),
        fillOpacity,
        strokeOpacity,
      });
    },
  );
  for (const category of Object.keys(categories)) {
    if (!additions.some((addition) => addition.category === category)) {
      fail("invalid-request", "Content append category state is unused.");
    }
  }
  const reusedResources = plan.reusedResources.map((item) => {
    if (!CONTENT_APPEND_CATEGORIES.includes(item.category)) {
      fail("invalid-request", "Content append reused category is invalid.");
    }
    claim(item.category, item.resourceName);
    assertObjectReference(
      item.reference.objectNumber,
      item.reference.generation,
    );
    return Object.freeze({
      category: item.category,
      resourceName: item.resourceName,
      reference: Object.freeze({
        objectNumber: item.reference.objectNumber,
        generation: item.reference.generation,
      }),
      expectedObjectSha256: exactSha256(
        item.expectedObjectSha256,
        "content append reused resource",
      ),
    });
  });
  if (
    plan.direction === "remove" &&
    (additions.length > 0 || reusedResources.length > 0)
  ) {
    fail("invalid-request", "Content removal cannot install or reuse resources.");
  }
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_CONTENT_APPEND_SCHEMA_VERSION,
    direction: plan.direction,
    appendKind: plan.appendKind,
    pageNumber,
    page: Object.freeze({
      objectNumber: plan.page.objectNumber,
      generation: plan.page.generation,
      expectedObjectSha256: exactSha256(
        plan.page.expectedObjectSha256,
        "content append page object",
      ),
    }),
    resourceOwner: Object.freeze({
      objectNumber: plan.resourceOwner.objectNumber,
      generation: plan.resourceOwner.generation,
      expectedObjectSha256: exactSha256(
        plan.resourceOwner.expectedObjectSha256,
        "content append resource owner",
      ),
      dictionaryKind: plan.resourceOwner.dictionaryKind,
    }),
    categories: Object.freeze(categories),
    additions: Object.freeze(additions),
    reusedResources: Object.freeze(reusedResources),
    restoredDepth: plan.restoredDepth,
  });
}

/** The replacement a request's pieces put together from its span (mirror of
 * the server's `_record_replacement`). */
export function expandPdfNativeReplacementPieces(
  span: Uint8Array,
  pieces: readonly PdfNativeReplacementPiece[],
): Uint8Array {
  return concatBytes(
    pieces.map((piece) =>
      "literal" in piece ? piece.literal : span.slice(piece.spanStart, piece.spanEnd),
    ),
  );
}

/** The duplicate spans of a request, in stream order: each non-empty, past
 * the one before it, and clear of the operator span. */
/**
 * Another span of the stream written with the edit. A whole copy holds the
 * operator span's bytes and takes the whole replacement. A partial copy (a
 * fake-bold overprint of one text object of a run group) holds the bytes of
 * `sourceStart..sourceEnd` inside the operator span and takes the part
 * `replacementStart..replacementEnd` of the replacement written over them.
 */
export type PdfNativeDuplicateSpan = Readonly<{
  start: number;
  end: number;
  sourceStart?: number;
  sourceEnd?: number;
  replacementStart?: number;
  replacementEnd?: number;
}>;

/** Each rule span lies clear of the operator span and the others, in order;
 * its bytes are checked against the stream, and its place against the text
 * object, where the candidate is built (mirror of replay). */
function normalizeTextRules(
  rules: readonly PdfNativeTextRuleSpan[],
  start: number,
  end: number,
): readonly PdfNativeTextRuleSpan[] {
  if (rules.length === 0 || rules.length > 2) {
    fail("invalid-request", "Text rule count is unsupported.");
  }
  const styles = new Set<string>();
  let cursor = 0;
  const out = [...rules].sort((left, right) => left.start - right.start).map((rule) => {
    if (
      (rule.style !== "underline" && rule.style !== "strikethrough") ||
      styles.has(rule.style) ||
      !Number.isSafeInteger(rule.start) || !Number.isSafeInteger(rule.end) ||
      rule.start < cursor || rule.end < rule.start ||
      (rule.start < end && rule.end > start) ||
      !isUint8Array(rule.removed) || !isUint8Array(rule.added) ||
      rule.removed.length !== rule.end - rule.start ||
      (rule.removed.length === 0 && rule.added.length === 0) ||
      rule.added.length > MAX_REPLACEMENT_BYTES
    ) {
      fail("invalid-request", "Text rule spans must be ordered, disjoint rules beside the operator span.");
    }
    styles.add(rule.style);
    cursor = rule.end;
    return Object.freeze({
      style: rule.style, start: rule.start, end: rule.end,
      removed: Uint8Array.from(rule.removed), added: Uint8Array.from(rule.added),
    });
  });
  return Object.freeze(out);
}

function normalizeDuplicateSpans(
  spans: readonly PdfNativeDuplicateSpan[],
  start: number,
  end: number,
  replacementLength: number,
): readonly PdfNativeDuplicateSpan[] {
  if (spans.length === 0 || spans.length > 64) {
    fail("invalid-request", "Duplicate span count is unsupported.");
  }
  let cursor = 0;
  const out = spans.map((span) => {
    const partialFields = [span.sourceStart, span.sourceEnd, span.replacementStart, span.replacementEnd];
    const partial = partialFields.some((value) => value !== undefined);
    if (partial && !partialFields.every((value) => Number.isSafeInteger(value))) {
      fail("invalid-request", "A partial duplicate span names its source and replacement range.");
    }
    const sourceStart = partial ? span.sourceStart! : start;
    const sourceEnd = partial ? span.sourceEnd! : end;
    if (!Number.isSafeInteger(span.start) || !Number.isSafeInteger(span.end) ||
      span.start < cursor || span.end <= span.start ||
      sourceStart < start || sourceEnd > end || sourceEnd <= sourceStart ||
      span.end - span.start !== sourceEnd - sourceStart ||
      (span.start < end && span.end > start) ||
      (partial && !(0 <= span.replacementStart! && span.replacementStart! <= span.replacementEnd! &&
        span.replacementEnd! <= replacementLength))) {
      fail("invalid-request", "Duplicate spans must be ordered, disjoint copies of the operator span.");
    }
    cursor = span.end;
    return Object.freeze(partial
      ? { start: span.start, end: span.end, sourceStart, sourceEnd,
          replacementStart: span.replacementStart!, replacementEnd: span.replacementEnd! }
      : { start: span.start, end: span.end });
  });
  return Object.freeze(out);
}

function normalizeReplacementPieces(
  pieces: readonly PdfNativeReplacementPiece[],
  spanLength: number,
): readonly PdfNativeReplacementPiece[] {
  if (
    !Array.isArray(pieces) ||
    pieces.length < 1 ||
    pieces.length > MAX_REPLACEMENT_PIECES
  ) {
    fail("invalid-request", "Replacement pieces are invalid.");
  }
  let literalBytes = 0;
  const ranges: [number, number][] = [];
  const normalized = pieces.map((piece) => {
    if ("literal" in piece) {
      if (!isUint8Array(piece.literal) || piece.literal.length === 0) {
        fail("invalid-request", "Replacement piece bytes are invalid.");
      }
      literalBytes += piece.literal.length;
      return Object.freeze({ literal: Uint8Array.from(piece.literal) });
    }
    const { spanStart, spanEnd } = piece;
    if (
      !Number.isSafeInteger(spanStart) ||
      !Number.isSafeInteger(spanEnd) ||
      spanStart < 0 ||
      spanEnd <= spanStart ||
      spanEnd > spanLength
    ) {
      fail("invalid-request", "Replacement piece range is invalid.");
    }
    ranges.push([spanStart, spanEnd]);
    return Object.freeze({ spanStart, spanEnd });
  });
  ranges.sort((left, right) => left[0] - right[0]);
  if (
    literalBytes > MAX_REPLACEMENT_BYTES ||
    ranges.some((range, index) => index > 0 && range[0] < ranges[index - 1][1])
  ) {
    fail("invalid-request", "Replacement pieces are invalid or too large.");
  }
  return Object.freeze(normalized);
}

function normalizeRequest(
  request: PdfNativeContentStreamTransactionRequest,
): PdfNativeContentStreamTransactionRequest {
  // A page font reference rides one plain write of the page's own text and
  // nothing beside it: no other change of the page's resources (a font
  // program installed in the same write among them), no joined page, no
  // pieces, copies or rules. Said before any of those is read.
  if (
    request.pageFontReference &&
    (request.pageExtGStateMutation || request.pageXObjectUnmap || request.pageFontExtensionMutation ||
      request.formInstanceIsolation || request.contentAppend || request.pageContentConsolidation ||
      request.replacementPieces || request.duplicateSpans || request.textRules)
  ) {
    fail("invalid-request", "A page font reference rides one plain write of the page's own text.");
  }
  assertObjectReference(
    request.identity.objectNumber,
    request.identity.generation,
  );
  const sourceSha256 = exactSha256(request.identity.sourceSha256, "source");
  const expectedObjectSha256 = exactSha256(
    request.expectedObjectSha256,
    "object precondition",
  );
  const expectedStreamSha256 = exactSha256(
    request.expectedStreamSha256,
    "stream precondition",
  );
  const expectedSpanSha256 = exactSha256(
    request.operatorSpan.expectedSha256,
    "operator span precondition",
  );
  const start = exactPositiveInteger(
    request.operatorSpan.start,
    "operator span start",
    true,
  );
  const end = exactPositiveInteger(
    request.operatorSpan.end,
    "operator span end",
    true,
  );
  const contentAppend = normalizeContentAppendPlan(request.contentAppend);
  const structActualText = normalizeStructActualText(request);
  const pageContentConsolidation = (() => {
    const plan = request.pageContentConsolidation;
    if (!plan) return undefined;
    if (
      !Number.isSafeInteger(plan.pageNumber) ||
      plan.pageNumber < 1 ||
      plan.pageNumber > 100_000 ||
      plan.sourceStreams.length < 1 ||
      plan.sourceStreams.length > 64 ||
      plan.stream.generation !== 0
    ) {
      fail("invalid-request", "Page content consolidation plan is invalid.");
    }
    assertObjectReference(plan.page.objectNumber, plan.page.generation);
    assertObjectReference(plan.stream.objectNumber, plan.stream.generation);
    const streamKeys = new Set<string>();
    for (const item of plan.sourceStreams) {
      assertObjectReference(item.objectNumber, item.generation);
      const key = `${item.objectNumber}:${item.generation}`;
      if (streamKeys.has(key) || item.objectNumber === plan.stream.objectNumber) {
        fail("invalid-request", "Page content consolidation streams are invalid.");
      }
      streamKeys.add(key);
    }
    return Object.freeze({
      pageNumber: plan.pageNumber,
      page: Object.freeze({
        objectNumber: plan.page.objectNumber,
        generation: plan.page.generation,
        expectedObjectSha256: exactSha256(plan.page.expectedObjectSha256, "consolidated page"),
      }),
      sourceStreams: Object.freeze(
        plan.sourceStreams.map((item) =>
          Object.freeze({
            objectNumber: item.objectNumber,
            generation: item.generation,
            expectedObjectSha256: exactSha256(item.expectedObjectSha256, "consolidated stream"),
          }),
        ),
      ),
      joinedDecodedSha256: exactSha256(plan.joinedDecodedSha256, "joined page content"),
      stream: Object.freeze({ objectNumber: plan.stream.objectNumber, generation: 0 as const }),
    });
  })();
  if (contentAppend?.direction === "append") {
    if (end !== start) {
      fail("invalid-request", "Content append span must be empty.");
    }
  } else if (end <= start) {
    fail("invalid-request", "Operator span must be non-empty.");
  }
  const replacementPieces = request.replacementPieces
    ? normalizeReplacementPieces(request.replacementPieces, end - start)
    : null;
  const duplicateSpans = request.duplicateSpans
    ? normalizeDuplicateSpans(request.duplicateSpans, start, end, request.replacement?.length ?? 0)
    : null;
  const textRules = request.textRules ? normalizeTextRules(request.textRules, start, end) : null;
  if (textRules && (duplicateSpans || request.formInstanceIsolation)) {
    fail("invalid-request", "Text rules are rewritten with one plain text replacement.");
  }
  // A box mark wraps one plain write of one span of the page's own stream
  // (mirror of the record contract's _box_mark_alone).
  const boxMark = request.boxMark
    ? (() => {
        const mark = request.boxMark;
        if (
          !Number.isSafeInteger(mark.start) || !Number.isSafeInteger(mark.end) ||
          mark.start < 0 || mark.end <= mark.start || mark.start > start || end > mark.end ||
          replacementPieces || duplicateSpans || textRules || request.formInstanceIsolation ||
          request.pageExtGStateMutation || request.pageXObjectUnmap || request.pageFontExtensionMutation ||
          request.pageFontReference ||
          request.contentAppend || request.pageContentConsolidation || request.structActualText
        ) {
          fail("invalid-request", "A box mark wraps one plain write round its span.");
        }
        return Object.freeze({ start: mark.start, end: mark.end });
      })()
    : null;
  // Pieces use each byte of the span at most once, so a replacement written
  // as pieces is no longer than the span and the bytes it adds.
  if (
    !isUint8Array(request.replacement) ||
    request.replacement.length >
      (replacementPieces ? end - start + MAX_REPLACEMENT_BYTES : MAX_REPLACEMENT_BYTES)
  ) {
    fail("invalid-request", "Replacement bytes are invalid or too large.");
  }
  const seen = new Set<string>();
  const untouched = (request.untouchedObjects ?? []).map((item) => {
    assertObjectReference(item.objectNumber, item.generation);
    if (
      item.objectNumber === request.identity.objectNumber &&
      item.generation === request.identity.generation
    ) {
      fail("invalid-request", "Target object cannot be marked untouched.");
    }
    const key = `${item.objectNumber}:${item.generation}`;
    if (seen.has(key))
      fail("invalid-request", "Untouched object is duplicated.");
    seen.add(key);
    return Object.freeze({
      objectNumber: item.objectNumber,
      generation: item.generation,
      expectedObjectSha256: exactSha256(
        item.expectedObjectSha256,
        "untouched object",
      ),
      role: item.role,
    });
  });
  const formInstanceIsolation = (() => {
    const isolation = request.formInstanceIsolation;
    if (!isolation) return undefined;
    assertObjectReference(
      isolation.page.objectNumber,
      isolation.page.generation,
    );
    if (isolation.chain.length < 1 || isolation.chain.length > 8) {
      fail(
        "invalid-request",
        "Form instance isolation chain is empty or too deep.",
      );
    }
    const chainSeen = new Set<string>();
    const pageKey = `${isolation.page.objectNumber}:${isolation.page.generation}`;
    if (
      seen.has(pageKey) ||
      (isolation.page.objectNumber === request.identity.objectNumber &&
        isolation.page.generation === request.identity.generation)
    ) {
      fail(
        "invalid-request",
        "A revised Form isolation page cannot be marked untouched or be the target.",
      );
    }
    const chain = isolation.chain.map((item) => {
      assertObjectReference(item.objectNumber, item.generation);
      if (!PDF_RESOURCE_NAME.test(item.resourceName)) {
        fail("invalid-request", "Form resource identity is invalid.");
      }
      const key = `${item.objectNumber}:${item.generation}`;
      if (key === pageKey || chainSeen.has(key)) {
        fail("invalid-request", "Form instance isolation chain is cyclic.");
      }
      chainSeen.add(key);
      if (item.kind !== undefined && item.kind !== "xobject" && item.kind !== "resources") {
        fail("invalid-request", "Form instance isolation link kind is invalid.");
      }
      if (item.kind === "resources" && item !== isolation.chain[0]) {
        fail("invalid-request", "Only the first isolation link may be a resources dictionary.");
      }
      const instance = item.instance;
      if (instance !== undefined) {
        if (
          item.kind === "resources" ||
          !instance ||
          typeof instance !== "object" ||
          !instance.contentStream ||
          typeof instance.contentStream !== "object" ||
          !PDF_RESOURCE_NAME.test(instance.resourceName) ||
          instance.resourceName === item.resourceName ||
          !Number.isSafeInteger(instance.operandStart) ||
          instance.operandStart < 0
        ) {
          fail("invalid-request", "Form instance isolation drawing is invalid.");
        }
        assertObjectReference(
          instance.contentStream.objectNumber,
          instance.contentStream.generation,
        );
        const streamKey = `${instance.contentStream.objectNumber}:${instance.contentStream.generation}`;
        if (seen.has(streamKey) || streamKey === pageKey) {
          fail("invalid-request", "A renamed Form drawing cannot sit in an untouched object.");
        }
      }
      return Object.freeze({
        resourceName: item.resourceName,
        kind: (item.kind ?? "xobject") as "xobject" | "resources",
        objectNumber: item.objectNumber,
        generation: item.generation,
        expectedObjectSha256: exactSha256(
          item.expectedObjectSha256,
          "Form object",
        ),
        ...(instance !== undefined
          ? {
              instance: Object.freeze({
                contentStream: Object.freeze({
                  objectNumber: instance.contentStream.objectNumber,
                  generation: instance.contentStream.generation,
                }),
                operandStart: instance.operandStart,
                resourceName: instance.resourceName,
              }),
            }
          : {}),
      });
    });
    const leaf = chain.at(-1)!;
    if (
      leaf.objectNumber !== request.identity.objectNumber ||
      leaf.generation !== request.identity.generation ||
      leaf.expectedObjectSha256 !== expectedObjectSha256
    ) {
      fail(
        "invalid-request",
        "Form isolation leaf does not match the selected target.",
      );
    }
    return Object.freeze({
      page: Object.freeze({
        objectNumber: isolation.page.objectNumber,
        generation: isolation.page.generation,
        expectedObjectSha256: exactSha256(
          isolation.page.expectedObjectSha256,
          "Form isolation page",
        ),
      }),
      chain: Object.freeze(chain),
    });
  })();
  const pageExtGStateMutation = (() => {
    const mutation = request.pageExtGStateMutation;
    if (!mutation) return undefined;
    if (formInstanceIsolation) {
      fail(
        "invalid-request",
        "Form isolation and page ExtGState mutation cannot share one transaction.",
      );
    }
    if (
      !Number.isSafeInteger(mutation.pageNumber) ||
      mutation.pageNumber < 1 ||
      mutation.pageNumber > 100_000 ||
      !PDF_RESOURCE_NAME.test(mutation.resourceName) ||
      typeof mutation.reuseExisting !== "boolean"
    ) {
      fail("invalid-request", "Page ExtGState mutation identity is invalid.");
    }
    assertObjectReference(mutation.page.objectNumber, mutation.page.generation);
    assertObjectReference(
      mutation.resourceOwner.objectNumber,
      mutation.resourceOwner.generation,
    );
    const pageKey = `${mutation.page.objectNumber}:${mutation.page.generation}`;
    if (
      pageKey ===
        `${request.identity.objectNumber}:${request.identity.generation}` ||
      seen.has(pageKey)
    ) {
      fail(
        "invalid-request",
        "A revised ExtGState page cannot be the stream target or untouched.",
      );
    }
    if (
      mutation.resourceOwner.dictionaryKind !== "nested-resources" &&
      mutation.resourceOwner.dictionaryKind !== "resource-dictionary"
    ) {
      fail("invalid-request", "Page Resources dictionary kind is invalid.");
    }
    const category = (() => {
      const expectedDictionarySha256 = exactSha256(
        mutation.category.expectedDictionarySha256,
        "ExtGState dictionary",
      );
      if (
        mutation.category.kind === "missing" ||
        mutation.category.kind === "direct"
      ) {
        return Object.freeze({
          kind: mutation.category.kind,
          expectedDictionarySha256,
        });
      }
      if (mutation.category.kind !== "indirect") {
        fail("invalid-request", "ExtGState category kind is invalid.");
      }
      assertObjectReference(
        mutation.category.target.objectNumber,
        mutation.category.target.generation,
      );
      return Object.freeze({
        kind: "indirect" as const,
        expectedDictionarySha256,
        target: Object.freeze({
          objectNumber: mutation.category.target.objectNumber,
          generation: mutation.category.target.generation,
          expectedObjectSha256: exactSha256(
            mutation.category.target.expectedObjectSha256,
            "ExtGState category object",
          ),
        }),
      });
    })();
    const fillOpacity = exactOpacity(mutation.fillOpacity, "Fill");
    const strokeOpacity = exactOpacity(mutation.strokeOpacity, "Stroke");
    const blendMode = exactRequestedBlendMode(mutation.blendMode);
    if (
      fillOpacity === null &&
      strokeOpacity === null &&
      blendMode === null
    ) {
      fail("invalid-request", "Page ExtGState appearance change is empty.");
    }
    return Object.freeze({
      pageNumber: mutation.pageNumber,
      page: Object.freeze({
        objectNumber: mutation.page.objectNumber,
        generation: mutation.page.generation,
        expectedObjectSha256: exactSha256(
          mutation.page.expectedObjectSha256,
          "ExtGState page object",
        ),
      }),
      resourceOwner: Object.freeze({
        objectNumber: mutation.resourceOwner.objectNumber,
        generation: mutation.resourceOwner.generation,
        expectedObjectSha256: exactSha256(
          mutation.resourceOwner.expectedObjectSha256,
          "ExtGState resource owner",
        ),
        dictionaryKind: mutation.resourceOwner.dictionaryKind,
      }),
      category,
      resourceName: mutation.resourceName,
      fillOpacity,
      strokeOpacity,
      blendMode,
      reuseExisting: mutation.reuseExisting,
    });
  })();
  if (
    contentAppend &&
    (formInstanceIsolation ||
      pageExtGStateMutation ||
      (contentAppend.direction === "remove") !==
        (request.replacement.length === 0))
  ) {
    fail(
      "invalid-request",
      "Content append cannot share another mutation or an inconsistent replacement.",
    );
  }
  const pageXObjectUnmap = (() => {
    const unmap = request.pageXObjectUnmap;
    if (!unmap) return undefined;
    if (formInstanceIsolation || pageExtGStateMutation) {
      fail(
        "invalid-request",
        "A page XObject unmap cannot ride with another resource mutation.",
      );
    }
    assertObjectReference(unmap.page.objectNumber, unmap.page.generation);
    assertObjectReference(
      unmap.resourceOwner.objectNumber,
      unmap.resourceOwner.generation,
    );
    if (
      !PDF_RESOURCE_NAME.test(unmap.resourceName) ||
      (unmap.direction !== "unmap" && unmap.direction !== "remap") ||
      !isUint8Array(unmap.ownerBytes) ||
      unmap.ownerBytes.length < 4 ||
      unmap.ownerBytes.length > MAX_REPLACEMENT_BYTES ||
      seen.has(
        `${unmap.resourceOwner.objectNumber}:${unmap.resourceOwner.generation}`,
      ) ||
      (unmap.resourceOwner.objectNumber === request.identity.objectNumber &&
        unmap.resourceOwner.generation === request.identity.generation)
    ) {
      fail("invalid-request", "Page XObject unmap authority is invalid.");
    }
    return Object.freeze({
      pageNumber: exactPositiveInteger(unmap.pageNumber, "unmap page number"),
      page: Object.freeze({ ...unmap.page }),
      resourceOwner: Object.freeze({ ...unmap.resourceOwner }),
      resourceName: unmap.resourceName,
      direction: unmap.direction,
      expectedOwnerSha256: exactSha256(
        unmap.expectedOwnerSha256,
        "page resource owner",
      ),
      ownerBytes: Uint8Array.from(unmap.ownerBytes),
      revisedOwnerSha256: exactSha256(
        unmap.revisedOwnerSha256,
        "revised page resource owner",
      ),
      entryOffset: exactPositiveInteger(unmap.entryOffset, "entry offset", true),
      entryBytes: Uint8Array.from(unmap.entryBytes),
    });
  })();
  const pageFontExtensionMutation = (() => {
    const mutation = request.pageFontExtensionMutation;
    if (!mutation) return undefined;
    // A font for a form's own resources goes into the edited form itself,
    // or into its clone when the edit isolates a shared form; a page font
    // never shares a transaction with an isolation.
    if (pageExtGStateMutation || (!mutation.form && formInstanceIsolation)) {
      fail(
        "invalid-request",
        "Font extension cannot share another resource mutation transaction.",
      );
    }
    if (
      !Number.isSafeInteger(mutation.pageNumber) ||
      mutation.pageNumber < 1 ||
      mutation.pageNumber > 100_000
    ) {
      fail("invalid-request", "Page font extension identity is invalid.");
    }
    assertObjectReference(mutation.page.objectNumber, mutation.page.generation);
    assertObjectReference(
      mutation.resourceOwner.objectNumber,
      mutation.resourceOwner.generation,
    );
    if (
      mutation.resourceOwner.dictionaryKind !== "nested-resources" &&
      mutation.resourceOwner.dictionaryKind !== "resource-dictionary"
    ) {
      fail("invalid-request", "Page font Resources kind is invalid.");
    }
    const category = (() => {
      const expectedDictionarySha256 = exactSha256(
        mutation.category.expectedDictionarySha256,
        "Font dictionary",
      );
      if (
        mutation.category.kind === "missing" ||
        mutation.category.kind === "direct"
      ) {
        return Object.freeze({
          kind: mutation.category.kind,
          expectedDictionarySha256,
        });
      }
      if (mutation.category.kind !== "indirect") {
        fail("invalid-request", "Font category kind is invalid.");
      }
      assertObjectReference(
        mutation.category.target.objectNumber,
        mutation.category.target.generation,
      );
      return Object.freeze({
        kind: "indirect" as const,
        expectedDictionarySha256,
        target: Object.freeze({
          objectNumber: mutation.category.target.objectNumber,
          generation: mutation.category.target.generation,
          expectedObjectSha256: exactSha256(
            mutation.category.target.expectedObjectSha256,
            "Font category object",
          ),
        }),
      });
    })();
    if (contentAppend) {
      fail(
        "invalid-request",
        "Content append cannot share a font extension transaction.",
      );
    }
    const fontPackage = normalizePageFontPackage(mutation.fontPackage);
    const additionalFontPackages = normalizeAdditionalFontPackages(fontPackage, mutation.additionalFontPackages);
    const replacement = asLatin1(request.replacement);
    for (const installed of pdfNativePageFontPackages({ fontPackage, additionalFontPackages })) {
      const encodedSegments = [
        ...replacement.matchAll(
          new RegExp(
            `/${escapeRegex(installed.resourceName)}\\s+[-+]?(?:\\d+(?:\\.\\d*)?|\\.\\d+)\\s+Tf\\s+<([0-9A-F]+)>\\s+Tj`,
            "g",
          ),
        ),
      ].map((match) => match[1]);
      if (
        (encodedSegments.length > 0 || !pdfNativeParagraphBlockUsesFontPackage(replacement, installed)) &&
        (encodedSegments.length < 1 ||
          encodedSegments.join("") !== installed.encodedHex) &&
        !pdfNativePropertyPartsUseFontPackage(replacement, installed)
      ) {
        fail(
          "invalid-request",
          "Font extension replacement does not use its proven font package.",
        );
      }
    }
    return Object.freeze({
      ...(additionalFontPackages ? { additionalFontPackages } : {}),
      pageNumber: mutation.pageNumber,
      page: Object.freeze({
        objectNumber: mutation.page.objectNumber,
        generation: mutation.page.generation,
        expectedObjectSha256: exactSha256(
          mutation.page.expectedObjectSha256,
          "Font extension page object",
        ),
      }),
      resourceOwner: Object.freeze({
        objectNumber: mutation.resourceOwner.objectNumber,
        generation: mutation.resourceOwner.generation,
        expectedObjectSha256: exactSha256(
          mutation.resourceOwner.expectedObjectSha256,
          "Font extension resource owner",
        ),
        dictionaryKind: mutation.resourceOwner.dictionaryKind,
      }),
      category,
      fontPackage,
      ...(mutation.form
        ? {
            form: (() => {
              assertObjectReference(
                mutation.form.objectNumber,
                mutation.form.generation,
              );
              return Object.freeze({
                objectNumber: mutation.form.objectNumber,
                generation: mutation.form.generation,
                expectedObjectSha256: exactSha256(
                  mutation.form.expectedObjectSha256,
                  "Font extension Form object",
                ),
              });
            })(),
          }
        : {}),
    });
  })();
  const pageFontReference = (() => {
    const reference = request.pageFontReference;
    if (!reference) return undefined;
    if (
      !Number.isSafeInteger(reference.pageNumber) || reference.pageNumber < 1 || reference.pageNumber > 100_000 ||
      !Array.isArray(reference.fonts) || reference.fonts.length < 1 ||
      reference.fonts.length > MAX_MRAI_NATIVE_PAGE_FONT_REFERENCES ||
      !Number.isSafeInteger(reference.resources?.objectNumber) || reference.resources.generation !== 0 ||
      (reference.resourceOwner.dictionaryKind !== "nested-resources" &&
        reference.resourceOwner.dictionaryKind !== "resource-dictionary") ||
      (reference.category.kind !== "missing" && reference.category.kind !== "direct" &&
        reference.category.kind !== "indirect")
    ) {
      fail("invalid-request", "Page font reference identity is invalid.");
    }
    assertObjectReference(reference.page.objectNumber, reference.page.generation);
    assertObjectReference(reference.resourceOwner.objectNumber, reference.resourceOwner.generation);
    assertObjectReference(reference.resources.objectNumber, 0);
    const drawn = asLatin1(request.replacement);
    const names = new Set<string>();
    const fonts = reference.fonts.map((item) => {
      assertObjectReference(item.font.objectNumber, item.font.generation);
      if (
        !PDF_RESOURCE_NAME.test(item.resourceName) || names.has(item.resourceName) ||
        // (Each font named is one the write draws in: nothing is named for nothing.)
        !new RegExp(`/${escapeRegex(item.resourceName)}[\\x00\\t\\n\\f\\r ]+[-+]?(?:\\d+(?:\\.\\d*)?|\\.\\d+)[\\x00\\t\\n\\f\\r ]+Tf(?![A-Za-z0-9*'"])`, "u").test(drawn)
      ) {
        fail("invalid-request", "A page font reference names fonts its write draws in, each once.");
      }
      names.add(item.resourceName);
      return Object.freeze({
        resourceName: item.resourceName,
        font: Object.freeze({
          objectNumber: item.font.objectNumber,
          generation: item.font.generation,
          expectedObjectSha256: exactSha256(item.font.expectedObjectSha256, "referenced font"),
        }),
      });
    });
    return Object.freeze({
      pageNumber: reference.pageNumber,
      page: Object.freeze({
        objectNumber: reference.page.objectNumber,
        generation: reference.page.generation,
        expectedObjectSha256: exactSha256(reference.page.expectedObjectSha256, "font reference page object"),
      }),
      resourceOwner: Object.freeze({
        objectNumber: reference.resourceOwner.objectNumber,
        generation: reference.resourceOwner.generation,
        expectedObjectSha256: exactSha256(reference.resourceOwner.expectedObjectSha256, "font reference resource owner"),
        dictionaryKind: reference.resourceOwner.dictionaryKind,
      }),
      category: reference.category.kind === "indirect"
        ? Object.freeze({
            kind: "indirect" as const,
            expectedDictionarySha256: exactSha256(reference.category.expectedDictionarySha256, "Font dictionary"),
            target: (() => {
              assertObjectReference(reference.category.target.objectNumber, reference.category.target.generation);
              return Object.freeze({
                objectNumber: reference.category.target.objectNumber,
                generation: reference.category.target.generation,
                expectedObjectSha256: exactSha256(reference.category.target.expectedObjectSha256, "Font category object"),
              });
            })(),
          })
        : Object.freeze({
            kind: reference.category.kind,
            expectedDictionarySha256: exactSha256(reference.category.expectedDictionarySha256, "Font dictionary"),
          }),
      resources: Object.freeze({ objectNumber: reference.resources.objectNumber, generation: 0 as const }),
      fonts: Object.freeze(fonts),
    });
  })();
  return Object.freeze({
    identity: Object.freeze({
      sourceSha256,
      objectNumber: request.identity.objectNumber,
      generation: request.identity.generation,
    }),
    expectedObjectSha256,
    expectedStreamSha256,
    operatorSpan: Object.freeze({
      start,
      end,
      expectedSha256: expectedSpanSha256,
    }),
    replacement: Uint8Array.from(request.replacement),
    ...(replacementPieces ? { replacementPieces } : {}),
    ...(duplicateSpans ? { duplicateSpans } : {}),
    ...(textRules ? { textRules } : {}),
    ...(boxMark ? { boxMark } : {}),
    untouchedObjects: Object.freeze(untouched),
    ...(formInstanceIsolation ? { formInstanceIsolation } : {}),
    ...(pageExtGStateMutation ? { pageExtGStateMutation } : {}),
    ...(pageXObjectUnmap ? { pageXObjectUnmap } : {}),
    ...(pageFontExtensionMutation ? { pageFontExtensionMutation } : {}),
    ...(pageFontReference ? { pageFontReference } : {}),
    ...(contentAppend ? { contentAppend } : {}),
    ...(pageContentConsolidation ? { pageContentConsolidation } : {}),
    ...(structActualText ? { structActualText } : {}),
  });
}

/** The one stream a reader sees for a page drawn from several: their
 * decoded bytes joined in order, stored as one new object with the
 * deterministic encoding every rewritten stream gets. Replay builds the
 * same object from the same streams. */
export async function buildMraiNativeConsolidatedContentStream(
  objectNumber: number,
  decoded: Uint8Array,
): Promise<
  Readonly<{
    bytes: Uint8Array;
    encodedStreamBytes: Uint8Array;
    objectSha256: string;
    streamSha256: string;
  }>
> {
  assertObjectReference(objectNumber, 0);
  const encodedStreamBytes = encodeDeterministicFlateStream(decoded);
  const bytes = concatBytes([
    ascii(
      `${objectNumber} 0 obj\n<< /Length ${encodedStreamBytes.length} /Filter /FlateDecode >>\nstream\n`,
    ),
    encodedStreamBytes,
    ascii("\nendstream\nendobj"),
  ]);
  return Object.freeze({
    bytes,
    encodedStreamBytes,
    objectSha256: await sha256PdfBytes(bytes),
    streamSha256: await sha256PdfBytes(encodedStreamBytes),
  });
}

/** The page's split streams joined, checked against the plan, and parsed as
 * the virtual target object the transaction edits. */
async function consolidatedVirtualTarget(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
  plan: PdfNativePageContentConsolidationPlan,
): Promise<Readonly<{ object: ParsedIndirectObject; bytes: Uint8Array }>> {
  const pages = await inspectMraiNativeDocumentPageGraphs(sourceBytes);
  const page = pages[plan.pageNumber - 1];
  if (
    !page ||
    page.pageNumber !== plan.pageNumber ||
    page.page.objectNumber !== plan.page.objectNumber ||
    page.page.generation !== plan.page.generation ||
    page.contentStreams.length !== plan.sourceStreams.length ||
    page.contentStreams.length < 1 ||
    page.contentStreams.some(
      (item, index) =>
        item.objectNumber !== plan.sourceStreams[index].objectNumber ||
        item.generation !== plan.sourceStreams[index].generation,
    )
  ) {
    fail("precondition-mismatch", "Page content stream set changed.");
  }
  const pageObject = parseIndirectObject(
    sourceBytes,
    parsed,
    plan.page.objectNumber,
    plan.page.generation,
  );
  if ((await sha256PdfBytes(pageObject.bytes)) !== plan.page.expectedObjectSha256) {
    fail("precondition-mismatch", "Page object precondition changed.");
  }
  const decoded: Uint8Array[] = [];
  for (const [index, item] of plan.sourceStreams.entries()) {
    const object = parseIndirectObject(
      sourceBytes,
      parsed,
      item.objectNumber,
      item.generation,
    );
    assertTargetStream(object);
    if ((await sha256PdfBytes(object.bytes)) !== item.expectedObjectSha256) {
      fail("precondition-mismatch", "Page content stream precondition changed.");
    }
    if (index > 0 && !hasSafeContentBoundaryBytes(decoded[index - 1], object.streamBytes)) {
      fail("unsupported-document", "PDF content-stream array splits a lexical token boundary.");
    }
    decoded.push(object.streamBytes);
  }
  const joined = concatBytes(decoded);
  if ((await sha256PdfBytes(joined)) !== plan.joinedDecodedSha256) {
    fail("precondition-mismatch", "Joined page content changed.");
  }
  const latest = latestEntries(parsed);
  if (
    plan.stream.generation !== 0 ||
    latest.get(plan.stream.objectNumber)?.inUse ||
    plan.stream.objectNumber < parsed.current.size
  ) {
    fail("invalid-request", "Consolidated content stream object number is in use.");
  }
  const virtual = await buildMraiNativeConsolidatedContentStream(
    plan.stream.objectNumber,
    joined,
  );
  const object = parseDirectObjectAtOffset(virtual.bytes, 0, {
    objectNumber: plan.stream.objectNumber,
    generation: 0,
  });
  assertTargetStream(object);
  return Object.freeze({ object, bytes: virtual.bytes });
}

/** The virtual joined stream of a consolidated page, inspected the way one
 * real stream is, for a writer that plans against the planner's proof. */
export async function inspectMraiNativeConsolidatedContentStream(
  sourceBytes: Uint8Array,
  plan: PdfNativePageContentConsolidationPlan,
): Promise<PdfNativeContentStreamInspection> {
  const parsed = parsePdf(sourceBytes, false);
  const sourceSha256 = await sha256PdfBytes(sourceBytes);
  const virtual = await consolidatedVirtualTarget(sourceBytes, parsed, plan);
  const object = virtual.object;
  assertTargetStream(object);
  return Object.freeze({
    identity: Object.freeze({
      sourceSha256,
      objectNumber: plan.stream.objectNumber,
      generation: 0,
    }),
    objectSha256: await sha256PdfBytes(virtual.bytes),
    streamSha256: await sha256PdfBytes(object.encodedStreamBytes),
    streamBytes: Uint8Array.from(object.streamBytes),
    untouchedObjects: Object.freeze([]),
  });
}

function hasSafeContentBoundaryBytes(left: Uint8Array, right: Uint8Array): boolean {
  if (left.length === 0 || right.length === 0) return false;
  const leftToken = String.fromCharCode(left[left.length - 1]);
  const rightToken = String.fromCharCode(right[0]);
  return (
    isPdfWhitespace(leftToken) ||
    isPdfWhitespace(rightToken) ||
    "()<>[]{}/%".includes(leftToken) ||
    "()<>[]{}/%".includes(rightToken)
  );
}

async function preparePageContentConsolidation(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
  request: PdfNativeContentStreamTransactionRequest,
  target: ParsedIndirectObject,
  virtualBytes: Uint8Array,
  changedDecodedStream: Uint8Array,
  changedStream: Uint8Array,
): Promise<
  Readonly<{
    candidateBytes: Uint8Array;
    changedObjectOffset: number;
    incrementalXrefOffset: number;
    changedObjectSha256: string;
    changedStreamSha256: string;
    changedDecodedStreamSha256: string;
    pageContentConsolidation: PdfNativePreparedPageContentConsolidation;
  }>
> {
  const plan = request.pageContentConsolidation;
  if (!plan) fail("invalid-request", "Page content consolidation is missing.");
  const revisedTarget = revisedContentStreamObject(virtualBytes, target, changedStream);
  const pageObject = parseIndirectObject(
    sourceBytes,
    parsed,
    plan.page.objectNumber,
    plan.page.generation,
  );
  // A page kept in an object stream is written out as its own object, the
  // way the font-extension rewrite does.
  const pageSource = asLatin1(pageObject.bytes);
  const revisedPageValue = latin1Bytes(
    rewriteDictionaryEntry(
      pageSource,
      "Contents",
      `${plan.stream.objectNumber} 0 R`,
      "PDF page Contents",
    ),
  );
  const revisedPage =
    pageObject.storage === "direct"
      ? revisedPageValue
      : concatBytes([
          ascii(`${pageObject.objectNumber} ${pageObject.generation} obj\n`),
          revisedPageValue,
          ascii("\nendobj"),
        ]);
  const revisions: IncrementalObjectRevision[] = [
    Object.freeze({
      objectNumber: plan.stream.objectNumber,
      generation: 0,
      bytes: revisedTarget,
    }),
    Object.freeze({
      objectNumber: plan.page.objectNumber,
      generation: plan.page.generation,
      bytes: revisedPage,
    }),
  ].sort(
    (left, right) =>
      left.objectNumber - right.objectNumber ||
      left.generation - right.generation,
  );
  const incremental = makeMultiObjectIncrementalRevision(
    sourceBytes,
    parsed.current,
    revisions,
    { min: 2, max: 2, label: "Page content consolidation" },
  );
  const preparedObject = async (revision: IncrementalObjectRevision) =>
    Object.freeze({
      objectNumber: revision.objectNumber,
      generation: revision.generation,
      offset: incremental.offsets.get(
        `${revision.objectNumber}:${revision.generation}`,
      )!,
      sha256: await sha256PdfBytes(revision.bytes),
    });
  const stream = await preparedObject(revisions.find((revision) => revision.objectNumber === plan.stream.objectNumber)!);
  const page = await preparedObject(revisions.find((revision) => revision.objectNumber === plan.page.objectNumber)!);
  return Object.freeze({
    candidateBytes: incremental.candidateBytes,
    changedObjectOffset: stream.offset,
    incrementalXrefOffset: incremental.incrementalXrefOffset,
    changedObjectSha256: stream.sha256,
    changedStreamSha256: await sha256PdfBytes(changedStream),
    changedDecodedStreamSha256: await sha256PdfBytes(changedDecodedStream),
    pageContentConsolidation: Object.freeze({ stream, page }),
  });
}

async function assertUntouchedPreconditions(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
  items: readonly PdfNativeObjectPrecondition[],
): Promise<void> {
  for (const item of items) {
    // Checking an object still stands is comparing the bytes the producer
    // wrote, never decoding them: the same reading `objectPrecondition` took
    // when it recorded the hash.
    const object = parseIndirectObject(
      sourceBytes,
      parsed,
      item.objectNumber,
      item.generation,
      undefined,
      "raw",
    );
    if ((await sha256PdfBytes(object.bytes)) !== item.expectedObjectSha256) {
      fail(
        "precondition-mismatch",
        `Untouched PDF object ${item.objectNumber} ${item.generation} changed.`,
      );
    }
  }
}

async function prepareFormInstanceIsolation(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
  request: PdfNativeContentStreamTransactionRequest,
  target: ParsedIndirectObject,
  changedDecodedStream: Uint8Array,
  changedStream: Uint8Array,
): Promise<
  Readonly<{
    candidateBytes: Uint8Array;
    changedObjectOffset: number;
    incrementalXrefOffset: number;
    changedObjectSha256: string;
    changedStreamSha256: string;
    changedDecodedStreamSha256: string;
    formInstanceIsolation: PdfNativePreparedFormIsolation;
    pageFontExtensionMutation?: PdfNativePreparedPageFontExtensionMutation;
  }>
> {
  const isolation = request.formInstanceIsolation;
  if (!isolation) {
    fail("invalid-request", "Form instance isolation plan is missing.");
  }
  if (target.storage !== "direct") {
    fail(
      "unsupported-document",
      "Selected Form must be a directly addressable stream object.",
    );
  }
  assertDirectlyWritableStream(target);
  const chainObjects: ParsedIndirectObject[] = [];
  for (const item of isolation.chain) {
    const parsedObject = parseIndirectObject(
      sourceBytes,
      parsed,
      item.objectNumber,
      item.generation,
    );
    const resourcesLink = item.kind === "resources";
    if (
      // A shared resources dictionary is copied, so it may sit in an object
      // stream; a Form is a stream object of its own.
      (resourcesLink
        ? parsedObject.storage !== "direct" && parsedObject.storage !== "object-stream"
        : parsedObject.storage !== "direct") ||
      (resourcesLink ? parsedObject.streamBytes !== null : parsedObject.streamBytes === null) ||
      (await sha256PdfBytes(parsedObject.bytes)) !== item.expectedObjectSha256
    ) {
      fail(
        "precondition-mismatch",
        "Form instance chain precondition changed.",
      );
    }
    const dictionary = asLatin1(parsedObject.bytes).split(
      /\bstream(?:\r\n|\r|\n)/,
      1,
    )[0];
    // A resources dictionary the pages share carries no /Subtype; a Form
    // carries /Subtype /Form and, optionally, /Type /XObject -- at its own
    // top level: the resources a Form holds inline name ExtGStates, fonts and
    // pictures with types of their own.
    const typeNames = [...dictionary.matchAll(/\/Type\s*\/([A-Za-z][A-Za-z0-9_.+\-]{0,63})/g)];
    const formNames = resourcesLink
      ? []
      : parseMraiNativeTopLevelDictionaryNameEntries(
          latin1Bytes(readMraiNativeObjectDictionarySource(sourceBytes, item) ?? ""),
        );
    const formTypes = formNames.filter((entry) => entry.key === "Type");
    const formSubtypes = formNames.filter((entry) => entry.key === "Subtype");
    if (
      resourcesLink
        ? typeNames.some((match) => match[1] !== "Resources") ||
          [...dictionary.matchAll(/\/Subtype\b/g)].length !== 0
        : formTypes.length > 1 ||
          (formTypes.length === 1 && formTypes[0].directName !== "XObject") ||
          formSubtypes.length !== 1 ||
          formSubtypes[0].directName !== "Form"
    ) {
      fail("unsupported-document", "Form instance chain identity is invalid.");
    }
    chainObjects.push(parsedObject);
  }
  const page = parseIndirectObject(
    sourceBytes,
    parsed,
    isolation.page.objectNumber,
    isolation.page.generation,
  );
  // A page kept in an object stream is written out as its own object, the
  // way the page content consolidation and the font extension do.
  if (
    (page.storage !== "direct" && page.storage !== "object-stream") ||
    page.streamBytes !== null ||
    (await sha256PdfBytes(page.bytes)) !== isolation.page.expectedObjectSha256
  ) {
    fail("precondition-mismatch", "Form isolation page precondition changed.");
  }

  const latest = latestEntries(parsed);
  const highestOwned = Math.max(0, ...latest.keys());
  const firstCloneNumber = Math.max(parsed.current.size, highestOwned + 1);
  if (firstCloneNumber + isolation.chain.length > 9_999_999_999) {
    fail("unsupported-document", "Form clone object number is too large.");
  }
  const cloneRefs = isolation.chain.map((_, index) =>
    Object.freeze({
      objectNumber: firstCloneNumber + (isolation.chain.length - 1 - index),
      generation: 0,
    }),
  );
  const leafIndex = isolation.chain.length - 1;
  let revisedLeaf = rewriteStreamObjectBytes(
    sourceBytes,
    target,
    changedStream,
  );
  // A font the edit needs goes into the cloned leaf's own resources: the
  // shared original Form keeps its dictionary, the clone names a new
  // Resources object that adds the font to what the form already had.
  const fontMutation = request.pageFontExtensionMutation ?? null;
  const fontRevisions: IncrementalObjectRevision[] = [];
  let fontResourcesRef: Readonly<{ objectNumber: number; generation: 0 }> | null = null;
  if (fontMutation) {
    const leaf = isolation.chain[leafIndex];
    if (
      !fontMutation.form ||
      fontMutation.form.objectNumber !== leaf.objectNumber ||
      fontMutation.form.generation !== leaf.generation ||
      fontMutation.form.expectedObjectSha256 !== leaf.expectedObjectSha256
    ) {
      fail(
        "invalid-request",
        "Form font extension must target the isolated leaf Form.",
      );
    }
    const pages = await inspectMraiNativeDocumentPageGraphs(sourceBytes);
    const owningPage = pages[fontMutation.pageNumber - 1];
    if (
      !owningPage ||
      owningPage.pageNumber !== fontMutation.pageNumber ||
      owningPage.page.objectNumber !== isolation.page.objectNumber ||
      owningPage.page.generation !== isolation.page.generation
    ) {
      fail("precondition-mismatch", "Form font extension page changed.");
    }
    const graph = await resolveFormFontGraph(sourceBytes, parsed, leaf);
    const rebuiltPlan = await formFontExtensionPlan(
      fontMutation.pageNumber,
      leaf,
      graph,
      fontMutation.fontPackage,
      fontMutation.additionalFontPackages,
    );
    if (
      !exactPageFontExtensionPlanMatches(fontMutation, rebuiltPlan) ||
      pdfNativePageFontPackages(fontMutation).some((installed) =>
        graph.entries.some((entry) => entry.resourceName === installed.resourceName) ||
        graph.names.has(installed.resourceName))
    ) {
      fail("precondition-mismatch", "Form Font resource authority changed.");
    }
    // The clones take the first free numbers; the font follows them.
    if (
      fontMutation.fontPackage.font.objectNumber !==
      firstCloneNumber + isolation.chain.length
    ) {
      fail("precondition-mismatch", "Font object allocation changed.");
    }
    await assertFontPackageIntegrity(sourceBytes, parsed, fontMutation);
    const resources = fontExtensionResourcesObject(graph, fontMutation);
    fontResourcesRef = resources.ref;
    fontRevisions.push(
      ...pageFontPackageObjects(fontMutation).map((object) =>
        Object.freeze({
          objectNumber: object.objectNumber,
          generation: object.generation,
          bytes: object.bytes,
        }),
      ),
      Object.freeze({ ...resources.ref, bytes: resources.bytes }),
    );
    revisedLeaf = rewriteStreamObjectResources(
      revisedLeaf,
      `${resources.ref.objectNumber} 0 R`,
    );
  }
  const leafClone = replaceIndirectObjectHeader(
    revisedLeaf,
    cloneRefs[leafIndex].objectNumber,
    0,
  );
  const revisions: IncrementalObjectRevision[] = [
    Object.freeze({ ...cloneRefs[leafIndex], bytes: leafClone }),
  ];
  // A drawing the page itself makes is renamed in the page's content
  // stream, which is written too.
  const pageStreamRevisions: IncrementalObjectRevision[] = [];
  const renamePageDrawing = async (
    link: (typeof isolation.chain)[number],
    instance: PdfNativeFormInstanceIsolationInstance,
  ) => {
    if (pageStreamRevisions.length > 0) {
      fail("unsupported-document", "One edit renames one drawing of the page.");
    }
    const pages = await inspectMraiNativeDocumentPageGraphs(sourceBytes);
    const owningPage = pages.find(
      (candidate) =>
        candidate.page.objectNumber === isolation.page.objectNumber &&
        candidate.page.generation === isolation.page.generation,
    );
    if (
      !owningPage ||
      !owningPage.contentStreams.some(
        (stream) =>
          stream.objectNumber === instance.contentStream.objectNumber &&
          stream.generation === instance.contentStream.generation,
      ) ||
      pages.reduce(
        (count, candidate) =>
          count +
          candidate.contentStreams.filter(
            (stream) =>
              stream.objectNumber === instance.contentStream.objectNumber &&
              stream.generation === instance.contentStream.generation,
          ).length,
        0,
      ) !== 1 ||
      countMraiNativeIndirectObjectReferences(sourceBytes, instance.contentStream) !== 1
    ) {
      fail("precondition-mismatch", "Renamed Form drawing is not this page's own stream.");
    }
    const stream = parseIndirectObject(
      sourceBytes,
      parsed,
      instance.contentStream.objectNumber,
      instance.contentStream.generation,
    );
    pageStreamRevisions.push(
      Object.freeze({
        objectNumber: instance.contentStream.objectNumber,
        generation: instance.contentStream.generation,
        bytes: renamedDrawingStreamBytes(sourceBytes, stream, link.resourceName, instance),
      }),
    );
  };
  for (let index = leafIndex - 1; index >= 0; index -= 1) {
    const child = isolation.chain[index + 1];
    const parentIsResources = isolation.chain[index].kind === "resources";
    let rebound: Uint8Array;
    if (child.instance) {
      // The parent draws the child more than once: its one drawing is
      // renamed and the parent's names gain the clone beside the original.
      if (parentIsResources) {
        await renamePageDrawing(child, child.instance);
        rebound = addDirectXObjectMapping(
          standaloneObjectBytes(chainObjects[index]),
          child.instance.resourceName,
          cloneRefs[index + 1],
          true,
        );
      } else {
        const parent = chainObjects[index];
        if (
          child.instance.contentStream.objectNumber !== parent.objectNumber ||
          child.instance.contentStream.generation !== parent.generation
        ) {
          fail("invalid-request", "A nested Form drawing is renamed in its parent Form.");
        }
        rebound = addDirectXObjectMapping(
          renamedDrawingStreamBytes(sourceBytes, parent, child.resourceName, child.instance),
          child.instance.resourceName,
          cloneRefs[index + 1],
          false,
        );
      }
    } else {
      rebound = replaceIsolationLink(
        standaloneObjectBytes(chainObjects[index]),
        child,
        child,
        cloneRefs[index + 1],
        parentIsResources,
      );
    }
    revisions.push(
      Object.freeze({
        ...cloneRefs[index],
        bytes: replaceIndirectObjectHeader(
          rebound,
          cloneRefs[index].objectNumber,
          0,
        ),
      }),
    );
  }
  const firstLink = isolation.chain[0];
  let revisedPage: Uint8Array;
  if (firstLink.instance) {
    await renamePageDrawing(firstLink, firstLink.instance);
    revisedPage = addDirectXObjectMapping(
      standaloneObjectBytes(page),
      firstLink.instance.resourceName,
      cloneRefs[0],
      false,
    );
  } else {
    revisedPage = replaceIsolationLink(
      standaloneObjectBytes(page),
      firstLink,
      firstLink,
      cloneRefs[0],
    );
  }
  revisions.push(
    Object.freeze({
      objectNumber: isolation.page.objectNumber,
      generation: isolation.page.generation,
      bytes: revisedPage,
    }),
    ...pageStreamRevisions,
    ...fontRevisions,
  );
  const incremental = makeMultiObjectIncrementalRevision(
    sourceBytes,
    parsed.current,
    revisions,
  );
  const objects = await Promise.all(
    revisions.map(async (revision) =>
      Object.freeze({
        objectNumber: revision.objectNumber,
        generation: revision.generation,
        offset: incremental.offsets.get(
          `${revision.objectNumber}:${revision.generation}`,
        )!,
        sha256: await sha256PdfBytes(revision.bytes),
      }),
    ),
  );
  const primaryOffset = incremental.offsets.get(
    `${cloneRefs[leafIndex].objectNumber}:0`,
  )!;
  return Object.freeze({
    candidateBytes: incremental.candidateBytes,
    changedObjectOffset: primaryOffset,
    incrementalXrefOffset: incremental.incrementalXrefOffset,
    changedObjectSha256: await sha256PdfBytes(leafClone),
    changedStreamSha256: await sha256PdfBytes(changedStream),
    changedDecodedStreamSha256: await sha256PdfBytes(changedDecodedStream),
    formInstanceIsolation: Object.freeze({
      primaryClone: cloneRefs[leafIndex],
      objects: Object.freeze(objects),
      clones: Object.freeze(
        isolation.chain
          .map((link, index) =>
            Object.freeze({
              original: Object.freeze({
                objectNumber: link.objectNumber,
                generation: link.generation,
              }),
              clone: cloneRefs[index],
            }),
          )
          .reverse(),
      ),
    }),
    ...(fontMutation && fontResourcesRef
      ? {
          pageFontExtensionMutation: Object.freeze({
            resourceName: fontMutation.fontPackage.resourceName,
            font: fontMutation.fontPackage.font,
            fontObjects: Object.freeze(
              pageFontPackageObjects(fontMutation).map(
                (object) =>
                  objects.find(
                    (item) =>
                      item.objectNumber === object.objectNumber &&
                      item.generation === object.generation,
                  )!,
              ),
            ),
            resources: objects.find(
              (item) => item.objectNumber === fontResourcesRef!.objectNumber,
            )!,
            page: objects.find(
              (item) => item.objectNumber === cloneRefs[leafIndex].objectNumber,
            )!,
          }),
        }
      : {}),
  });
}

function rewriteDictionaryEntry(
  source: string,
  key: string,
  value: string,
  label: string,
): string {
  const entries = topLevelDictionaryEntries(source);
  const matches = entries.filter((entry) => entry.key === key);
  if (matches.length > 1) {
    fail("unsupported-document", `${label} is ambiguous.`);
  }
  if (matches.length === 1) {
    const entry = matches[0];
    // A value that starts with a regular character (an indirect reference,
    // number or keyword) needs whitespace after a key it directly follows;
    // `/Resources<<` rewritten to `/Resources6610 0 R` would rename the key.
    const separator =
      !isPdfWhitespace(source[entry.valueStart - 1]) &&
      !/^[<\[(/]/.test(value)
        ? " "
        : "";
    return `${source.slice(0, entry.valueStart)}${separator}${value}${source.slice(entry.valueEnd)}`;
  }
  const start = source.indexOf("<<");
  const end = dictionaryEnd(source, start) - 2;
  let insertion = end;
  while (insertion > start + 2 && isPdfWhitespace(source[insertion - 1])) {
    insertion -= 1;
  }
  return `${source.slice(0, insertion)}\n/${key} ${value}\n${source.slice(end)}`;
}

function revisedContentStreamObject(
  sourceBytes: Uint8Array,
  target: ParsedIndirectObject,
  changedStream: Uint8Array,
): Uint8Array {
  assertDirectlyWritableStream(target);
  return rewriteStreamObjectBytes(sourceBytes, target, changedStream);
}

function exactPageExtGStatePlanMatches(
  left: PdfNativePageExtGStateMutationPlan,
  right: PdfNativePageExtGStateMutationPlan,
): boolean {
  return JSON.stringify(left) === JSON.stringify(right);
}

async function preparePageXObjectUnmap(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
  request: PdfNativeContentStreamTransactionRequest,
  target: ParsedIndirectObject,
  changedDecodedStream: Uint8Array,
  changedStream: Uint8Array,
): Promise<
  Readonly<{
    candidateBytes: Uint8Array;
    changedObjectOffset: number;
    incrementalXrefOffset: number;
    changedObjectSha256: string;
    changedStreamSha256: string;
    changedDecodedStreamSha256: string;
    pageXObjectUnmap: Readonly<{
      resourceName: string;
      direction: "unmap" | "remap";
      resourceOwner: PdfNativePreparedIsolatedObject;
    }>;
  }>
> {
  const unmap = request.pageXObjectUnmap;
  if (!unmap) fail("invalid-request", "Page XObject unmap is missing.");
  const owner = parseIndirectObject(
    sourceBytes,
    parsed,
    unmap.resourceOwner.objectNumber,
    unmap.resourceOwner.generation,
  );
  if (
    (await sha256PdfBytes(owner.bytes)) !== unmap.expectedOwnerSha256 ||
    (await sha256PdfBytes(unmap.ownerBytes)) !== unmap.revisedOwnerSha256 ||
    sameNativeObjectRef(unmap.resourceOwner, request.identity)
  ) {
    fail("precondition-mismatch", "PDF page resource authority changed.");
  }
  const revisedTarget = revisedContentStreamObject(
    sourceBytes,
    target,
    changedStream,
  );
  const revisions: IncrementalObjectRevision[] = [
    Object.freeze({
      objectNumber: request.identity.objectNumber,
      generation: request.identity.generation,
      bytes: revisedTarget,
    }),
    Object.freeze({
      objectNumber: unmap.resourceOwner.objectNumber,
      generation: unmap.resourceOwner.generation,
      bytes: Uint8Array.from(unmap.ownerBytes),
    }),
  ].sort(
    (left, right) =>
      left.objectNumber - right.objectNumber ||
      left.generation - right.generation,
  );
  const incremental = makeMultiObjectIncrementalRevision(
    sourceBytes,
    parsed.current,
    revisions,
    { min: 2, max: 2, label: "Page XObject unmap" },
  );
  const preparedObject = async (revision: IncrementalObjectRevision) =>
    Object.freeze({
      objectNumber: revision.objectNumber,
      generation: revision.generation,
      offset: incremental.offsets.get(
        `${revision.objectNumber}:${revision.generation}`,
      )!,
      sha256: await sha256PdfBytes(revision.bytes),
    });
  const targetPrepared = await preparedObject(
    revisions.find(
      (revision) =>
        revision.objectNumber === request.identity.objectNumber &&
        revision.generation === request.identity.generation,
    )!,
  );
  return Object.freeze({
    candidateBytes: incremental.candidateBytes,
    changedObjectOffset: targetPrepared.offset,
    incrementalXrefOffset: incremental.incrementalXrefOffset,
    changedObjectSha256: targetPrepared.sha256,
    changedStreamSha256: await sha256PdfBytes(changedStream),
    changedDecodedStreamSha256: await sha256PdfBytes(changedDecodedStream),
    pageXObjectUnmap: Object.freeze({
      resourceName: unmap.resourceName,
      direction: unmap.direction,
      resourceOwner: await preparedObject(
        revisions.find(
          (revision) =>
            revision.objectNumber === unmap.resourceOwner.objectNumber &&
            revision.generation === unmap.resourceOwner.generation,
        )!,
      ),
    }),
  });
}

async function preparePageExtGStateMutation(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
  request: PdfNativeContentStreamTransactionRequest,
  target: ParsedIndirectObject,
  changedDecodedStream: Uint8Array,
  changedStream: Uint8Array,
): Promise<
  Readonly<{
    candidateBytes: Uint8Array;
    changedObjectOffset: number;
    incrementalXrefOffset: number;
    changedObjectSha256: string;
    changedStreamSha256: string;
    changedDecodedStreamSha256: string;
    pageExtGStateMutation: PdfNativePreparedPageExtGStateMutation;
  }>
> {
  const mutation = request.pageExtGStateMutation;
  if (!mutation) fail("invalid-request", "Page ExtGState mutation is missing.");
  const pages = await inspectMraiNativeDocumentPageGraphs(sourceBytes);
  const page = pages[mutation.pageNumber - 1];
  // The state is installed on the page's own Resources, so the edit may sit
  // in any one of the streams the page draws from.
  if (
    !page ||
    page.pageNumber !== mutation.pageNumber ||
    page.contentStreams.filter((stream) =>
      sameNativeObjectRef(stream, request.identity),
    ).length !== 1
  ) {
    fail(
      "unsupported-document",
      "Page ExtGState target stream ownership is shared or ambiguous.",
    );
  }
  const graph = await resolvePageExtGStateGraph(sourceBytes, parsed, page);
  const existing = graph.inspection.entries.filter(
    (entry) => entry.resourceName === mutation.resourceName,
  );
  const rebuiltPlan = Object.freeze({
    pageNumber: graph.inspection.pageNumber,
    page: graph.inspection.page,
    resourceOwner: graph.inspection.resourceOwner,
    category: graph.inspection.category,
    resourceName: mutation.resourceName,
    fillOpacity: mutation.fillOpacity,
    strokeOpacity: mutation.strokeOpacity,
    blendMode: mutation.blendMode,
    reuseExisting: existing.length === 1,
  });
  if (
    !exactPageExtGStatePlanMatches(mutation, rebuiltPlan) ||
    (existing.length === 1 &&
      (!existing[0].supportedAppearanceOnly ||
        existing[0].fillOpacity !== mutation.fillOpacity ||
        existing[0].strokeOpacity !== mutation.strokeOpacity ||
        existing[0].blendMode !== mutation.blendMode))
  ) {
    fail("precondition-mismatch", "Page ExtGState authority changed.");
  }
  const revisedTarget = revisedContentStreamObject(
    sourceBytes,
    target,
    changedStream,
  );
  if (mutation.reuseExisting) {
    const incremental = makeIncrementalRevision(
      sourceBytes,
      parsed.current,
      request.identity.objectNumber,
      request.identity.generation,
      revisedTarget,
    );
    return Object.freeze({
      candidateBytes: incremental.candidateBytes,
      changedObjectOffset: incremental.changedObjectOffset,
      incrementalXrefOffset: incremental.incrementalXrefOffset,
      changedObjectSha256: await sha256PdfBytes(revisedTarget),
      changedStreamSha256: await sha256PdfBytes(changedStream),
      changedDecodedStreamSha256: await sha256PdfBytes(changedDecodedStream),
      pageExtGStateMutation: Object.freeze({
        resourceName: mutation.resourceName,
        reused: true,
      }),
    });
  }
  if (existing.length !== 0) {
    fail("unsupported-document", "Page ExtGState resource name is not unique.");
  }
  const latest = latestEntries(parsed);
  const firstNewObject = Math.max(
    parsed.current.size,
    Math.max(0, ...latest.keys()) + 1,
  );
  if (firstNewObject + 1 > 9_999_999_999) {
    fail("unsupported-document", "Page ExtGState object number is too large.");
  }
  const graphicsStateRef = Object.freeze({
    objectNumber: firstNewObject,
    generation: 0,
  });
  const resourcesRef = Object.freeze({
    objectNumber: firstNewObject + 1,
    generation: 0,
  });
  const appearanceEntries = [
    mutation.fillOpacity === null
      ? null
      : `/ca ${formatNativePdfNumber(mutation.fillOpacity)}`,
    mutation.strokeOpacity === null
      ? null
      : `/CA ${formatNativePdfNumber(mutation.strokeOpacity)}`,
    mutation.blendMode === null
      ? null
      : `/BM /${PDF_NAME_BY_BLEND_MODE[mutation.blendMode]}`,
  ].filter((entry): entry is string => entry !== null);
  const graphicsStateObject = ascii(
    `${graphicsStateRef.objectNumber} 0 obj\n<< /Type /ExtGState ${appearanceEntries.join(" ")} >>\nendobj`,
  );
  const categoryDictionary = rewriteDictionaryEntry(
    graph.categoryDictionary,
    mutation.resourceName,
    `${graphicsStateRef.objectNumber} 0 R`,
    "PDF ExtGState resource name",
  );
  const resourcesDictionary = rewriteDictionaryEntry(
    graph.resourcesDictionary,
    "ExtGState",
    categoryDictionary,
    "PDF ExtGState category",
  );
  const resourcesObject = latin1Bytes(
    `${resourcesRef.objectNumber} 0 obj\n${resourcesDictionary}\nendobj`,
  );
  const pageSource = asLatin1(graph.pageObject.bytes);
  const revisedPageValue = latin1Bytes(
    rewriteDictionaryEntry(
      pageSource,
      "Resources",
      `${resourcesRef.objectNumber} 0 R`,
      "PDF page Resources",
    ),
  );
  const revisedPage =
    graph.pageObject.storage === "direct"
      ? revisedPageValue
      : concatBytes([
          ascii(
            `${graph.pageObject.objectNumber} ${graph.pageObject.generation} obj\n`,
          ),
          revisedPageValue,
          ascii("\nendobj"),
        ]);
  const revisions: IncrementalObjectRevision[] = [
    Object.freeze({
      objectNumber: request.identity.objectNumber,
      generation: request.identity.generation,
      bytes: revisedTarget,
    }),
    Object.freeze({
      objectNumber: mutation.page.objectNumber,
      generation: mutation.page.generation,
      bytes: revisedPage,
    }),
    Object.freeze({ ...graphicsStateRef, bytes: graphicsStateObject }),
    Object.freeze({ ...resourcesRef, bytes: resourcesObject }),
  ].sort(
    (left, right) =>
      left.objectNumber - right.objectNumber ||
      left.generation - right.generation,
  );
  const incremental = makeMultiObjectIncrementalRevision(
    sourceBytes,
    parsed.current,
    revisions,
    { min: 4, max: 4, label: "Page ExtGState" },
  );
  const preparedObject = async (revision: IncrementalObjectRevision) =>
    Object.freeze({
      objectNumber: revision.objectNumber,
      generation: revision.generation,
      offset: incremental.offsets.get(
        `${revision.objectNumber}:${revision.generation}`,
      )!,
      sha256: await sha256PdfBytes(revision.bytes),
    });
  const targetPrepared = await preparedObject(
    revisions.find(
      (revision) =>
        revision.objectNumber === request.identity.objectNumber &&
        revision.generation === request.identity.generation,
    )!,
  );
  return Object.freeze({
    candidateBytes: incremental.candidateBytes,
    changedObjectOffset: targetPrepared.offset,
    incrementalXrefOffset: incremental.incrementalXrefOffset,
    changedObjectSha256: targetPrepared.sha256,
    changedStreamSha256: await sha256PdfBytes(changedStream),
    changedDecodedStreamSha256: await sha256PdfBytes(changedDecodedStream),
    pageExtGStateMutation: Object.freeze({
      resourceName: mutation.resourceName,
      reused: false,
      graphicsState: await preparedObject(
        revisions.find(
          (revision) => revision.objectNumber === graphicsStateRef.objectNumber,
        )!,
      ),
      resources: await preparedObject(
        revisions.find(
          (revision) => revision.objectNumber === resourcesRef.objectNumber,
        )!,
      ),
      page: await preparedObject(
        revisions.find(
          (revision) =>
            revision.objectNumber === mutation.page.objectNumber &&
            revision.generation === mutation.page.generation,
        )!,
      ),
    }),
  });
}

function exactPageFontExtensionPlanMatches(
  left: PdfNativePageFontExtensionMutationPlan,
  right: PdfNativePageFontExtensionMutationPlan,
): boolean {
  const summary = (value: PdfNativePageFontExtensionMutationPlan) => ({
    pageNumber: value.pageNumber,
    page: value.page,
    resourceOwner: value.resourceOwner,
    category: value.category,
    form: value.form ?? null,
    fontPackage: {
      schemaVersion: value.fontPackage.schemaVersion,
      fontSha256: value.fontPackage.fontSha256,
      resourceName: value.fontPackage.resourceName,
      font: value.fontPackage.font,
      ...(value.fontPackage.schemaVersion ===
      MRAI_NATIVE_FONT_ALIAS_PACKAGE_SCHEMA_VERSION
        ? {
            sourceFont: value.fontPackage.sourceFont,
            sourceToUnicode: value.fontPackage.sourceToUnicode,
          }
        : {
            unitsPerEm: value.fontPackage.unitsPerEm,
            ascent: value.fontPackage.ascent,
            descent: value.fontPackage.descent,
          }),
      encodedHex: value.fontPackage.encodedHex,
      widthAt1000: value.fontPackage.widthAt1000,
      objectByteLength: value.fontPackage.objectByteLength,
      objects: value.fontPackage.objects.map((object) => ({
        objectNumber: object.objectNumber,
        generation: object.generation,
        sha256: object.sha256,
        byteLength: object.bytes.length,
      })),
    },
    additionalFontPackages: (value.additionalFontPackages ?? []).map((item) => ({
      fontSha256: item.fontSha256,
      programFace: item.programFace ?? null,
      resourceName: item.resourceName,
      font: item.font,
      encodedHex: item.encodedHex,
      widthAt1000: item.widthAt1000,
      objects: item.objects.map((object) => ({
        objectNumber: object.objectNumber,
        sha256: object.sha256,
        byteLength: object.bytes.length,
      })),
    })),
  });
  return JSON.stringify(summary(left)) === JSON.stringify(summary(right));
}

/** The package bytes still hash as planned, and an alias package's source
 * font objects are still the ones it was built from. */
async function assertFontPackageIntegrity(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
  mutation: PdfNativePageFontExtensionMutationPlan,
): Promise<void> {
  for (const object of pageFontPackageObjects(mutation)) {
    if ((await sha256PdfBytes(object.bytes)) !== object.sha256) {
      fail("precondition-mismatch", "Font package object bytes changed.");
    }
  }
  if (
    mutation.fontPackage.schemaVersion ===
    MRAI_NATIVE_FONT_ALIAS_PACKAGE_SCHEMA_VERSION
  ) {
    for (const witness of [
      mutation.fontPackage.sourceFont,
      mutation.fontPackage.sourceToUnicode,
    ]) {
      const object = parseIndirectObject(
        sourceBytes,
        parsed,
        witness.objectNumber,
        witness.generation,
      );
      if ((await sha256PdfBytes(object.bytes)) !== witness.expectedObjectSha256) {
        fail("precondition-mismatch", "Font alias source object changed.");
      }
    }
  }
}

/** The Resources dictionary that names one new font, ready to be written as
 * its own object: the category rewritten, everything else as it was. */
function fontExtensionResourcesObject(
  graph: ResolvedPageFontGraph,
  mutation: PdfNativePageFontExtensionMutationPlan,
): Readonly<{ ref: Readonly<{ objectNumber: number; generation: 0 }>; bytes: Uint8Array }> {
  const ref = Object.freeze({
    objectNumber: pageFontPackageObjects(mutation).at(-1)!.objectNumber + 1,
    generation: 0 as const,
  });
  if (ref.objectNumber > 9_999_999_999) {
    fail("unsupported-document", "Font resource object number is too large.");
  }
  let categoryDictionary = graph.categoryDictionary;
  for (const installed of pdfNativePageFontPackages(mutation)) {
    categoryDictionary = rewriteDictionaryEntry(
      categoryDictionary,
      installed.resourceName,
      `${installed.font.objectNumber} 0 R`,
      "PDF Font resource name",
    );
  }
  const resourcesDictionary = rewriteDictionaryEntry(
    graph.resourcesDictionary,
    "Font",
    categoryDictionary,
    "PDF Font category",
  );
  return Object.freeze({
    ref,
    bytes: latin1Bytes(`${ref.objectNumber} 0 obj\n${resourcesDictionary}\nendobj`),
  });
}

/** A font for text an unshared Form XObject draws: the form's own stream
 * is rewritten and its dictionary names a new Resources object that adds
 * the font to what the form already had. */
async function prepareFormFontExtensionMutation(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
  request: PdfNativeContentStreamTransactionRequest,
  target: ParsedIndirectObject,
  changedDecodedStream: Uint8Array,
  changedStream: Uint8Array,
): Promise<
  Readonly<{
    candidateBytes: Uint8Array;
    changedObjectOffset: number;
    incrementalXrefOffset: number;
    changedObjectSha256: string;
    changedStreamSha256: string;
    changedDecodedStreamSha256: string;
    pageFontExtensionMutation: PdfNativePreparedPageFontExtensionMutation;
    /** Never set here: a form font extension refuses a consolidated page.
     * Declared so both font-extension preparers return one shape. */
    pageContentConsolidation?: PdfNativePreparedPageContentConsolidation;
  }>
> {
  const mutation = request.pageFontExtensionMutation;
  if (!mutation?.form) fail("invalid-request", "Form font extension is missing.");
  if (request.pageContentConsolidation) {
    fail("invalid-request", "Form font extension cannot ride on a consolidated page.");
  }
  if (
    mutation.form.objectNumber !== request.identity.objectNumber ||
    mutation.form.generation !== request.identity.generation
  ) {
    fail("invalid-request", "Form font extension must target the edited Form.");
  }
  const pages = await inspectMraiNativeDocumentPageGraphs(sourceBytes);
  const page = pages[mutation.pageNumber - 1];
  if (!page || page.pageNumber !== mutation.pageNumber) {
    fail("precondition-mismatch", "Form font extension page changed.");
  }
  const graph = await resolveFormFontGraph(sourceBytes, parsed, request.identity);
  const rebuiltPlan = await formFontExtensionPlan(
    mutation.pageNumber,
    request.identity,
    graph,
    mutation.fontPackage,
    mutation.additionalFontPackages,
  );
  if (
    !exactPageFontExtensionPlanMatches(mutation, rebuiltPlan) ||
    pdfNativePageFontPackages(mutation).some((installed) =>
        graph.entries.some((entry) => entry.resourceName === installed.resourceName) ||
        graph.names.has(installed.resourceName))
  ) {
    fail("precondition-mismatch", "Form Font resource authority changed.");
  }
  const latest = latestEntries(parsed);
  const firstNewObject = Math.max(
    parsed.current.size,
    Math.max(0, ...latest.keys()) + 1,
  );
  if (mutation.fontPackage.font.objectNumber !== firstNewObject) {
    fail("precondition-mismatch", "Font object allocation changed.");
  }
  await assertFontPackageIntegrity(sourceBytes, parsed, mutation);
  const resources = fontExtensionResourcesObject(graph, mutation);
  const revisedTarget = rewriteStreamObjectResources(
    revisedContentStreamObject(sourceBytes, target, changedStream),
    `${resources.ref.objectNumber} 0 R`,
  );
  const revisions: IncrementalObjectRevision[] = [
    Object.freeze({
      objectNumber: request.identity.objectNumber,
      generation: request.identity.generation,
      bytes: revisedTarget,
    }),
    ...pageFontPackageObjects(mutation).map((object) =>
      Object.freeze({
        objectNumber: object.objectNumber,
        generation: object.generation,
        bytes: object.bytes,
      }),
    ),
    Object.freeze({ ...resources.ref, bytes: resources.bytes }),
  ].sort(
    (left, right) =>
      left.objectNumber - right.objectNumber ||
      left.generation - right.generation,
  );
  const incremental = makeMultiObjectIncrementalRevision(
    sourceBytes,
    parsed.current,
    revisions,
    {
      min: pageFontPackageObjects(mutation).length + 2,
      max: pageFontPackageObjects(mutation).length + 2,
      label: "Form Font extension",
    },
  );
  const preparedObject = async (revision: IncrementalObjectRevision) =>
    Object.freeze({
      objectNumber: revision.objectNumber,
      generation: revision.generation,
      offset: incremental.offsets.get(
        `${revision.objectNumber}:${revision.generation}`,
      )!,
      sha256: await sha256PdfBytes(revision.bytes),
    });
  const find = (objectNumber: number, generation: number) =>
    revisions.find(
      (revision) =>
        revision.objectNumber === objectNumber &&
        revision.generation === generation,
    )!;
  const targetPrepared = await preparedObject(
    find(request.identity.objectNumber, request.identity.generation),
  );
  return Object.freeze({
    candidateBytes: incremental.candidateBytes,
    changedObjectOffset: targetPrepared.offset,
    incrementalXrefOffset: incremental.incrementalXrefOffset,
    changedObjectSha256: targetPrepared.sha256,
    changedStreamSha256: await sha256PdfBytes(changedStream),
    changedDecodedStreamSha256: await sha256PdfBytes(changedDecodedStream),
    pageFontExtensionMutation: Object.freeze({
      resourceName: mutation.fontPackage.resourceName,
      font: mutation.fontPackage.font,
      fontObjects: Object.freeze(
        await Promise.all(
          pageFontPackageObjects(mutation).map((object) =>
            preparedObject(find(object.objectNumber, object.generation)),
          ),
        ),
      ),
      resources: await preparedObject(find(resources.ref.objectNumber, 0)),
      page: targetPrepared,
    }),
  });
}

async function preparePageFontExtensionMutation(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
  request: PdfNativeContentStreamTransactionRequest,
  target: ParsedIndirectObject,
  changedDecodedStream: Uint8Array,
  changedStream: Uint8Array,
  consolidatedBytes: Uint8Array | null = null,
): Promise<
  Readonly<{
    candidateBytes: Uint8Array;
    changedObjectOffset: number;
    incrementalXrefOffset: number;
    changedObjectSha256: string;
    changedStreamSha256: string;
    changedDecodedStreamSha256: string;
    pageFontExtensionMutation: PdfNativePreparedPageFontExtensionMutation;
    pageContentConsolidation?: PdfNativePreparedPageContentConsolidation;
  }>
> {
  const mutation = request.pageFontExtensionMutation;
  if (!mutation) fail("invalid-request", "Page font extension is missing.");
  const pages = await inspectMraiNativeDocumentPageGraphs(sourceBytes);
  const page = pages[mutation.pageNumber - 1];
  // The rewritten stream has to be exactly one of this page's content
  // streams; the others stay pinned as untouched objects by the transaction.
  // A consolidated page edits its virtual joined stream, matched to the
  // page's streams already; the page then names it beside the new font.
  if (
    !page ||
    page.pageNumber !== mutation.pageNumber ||
    (!request.pageContentConsolidation &&
      page.contentStreams.filter(
        (candidate) =>
          candidate.objectNumber === request.identity.objectNumber &&
          candidate.generation === request.identity.generation,
      ).length !== 1)
  ) {
    fail(
      "unsupported-document",
      "Page font extension stream ownership is shared or ambiguous.",
    );
  }
  const graph = await resolvePageFontGraph(sourceBytes, parsed, page);
  const rebuiltPlan = Object.freeze({
    pageNumber: page.pageNumber,
    page: Object.freeze({
      ...page.page,
      expectedObjectSha256: await sha256PdfBytes(graph.pageObject.bytes),
    }),
    resourceOwner: Object.freeze({
      ...page.resourceAuthority.owner,
      expectedObjectSha256: await sha256PdfBytes(
        graph.resourceOwnerObject.bytes,
      ),
      dictionaryKind: page.resourceAuthority.dictionaryKind,
    }),
    category: graph.category,
    fontPackage: mutation.fontPackage,
    ...(mutation.additionalFontPackages ? { additionalFontPackages: mutation.additionalFontPackages } : {}),
  });
  if (
    !exactPageFontExtensionPlanMatches(mutation, rebuiltPlan) ||
    pdfNativePageFontPackages(mutation).some((installed) =>
        graph.entries.some((entry) => entry.resourceName === installed.resourceName) ||
        graph.names.has(installed.resourceName))
  ) {
    fail("precondition-mismatch", "Page Font resource authority changed.");
  }
  const latest = latestEntries(parsed);
  const firstNewObject = Math.max(
    parsed.current.size,
    Math.max(0, ...latest.keys()) + 1,
  );
  // The consolidated stream takes the first free number; the font follows.
  if (
    request.pageContentConsolidation &&
    request.pageContentConsolidation.stream.objectNumber !== firstNewObject
  ) {
    fail("precondition-mismatch", "Consolidated stream allocation changed.");
  }
  if (
    mutation.fontPackage.font.objectNumber !==
    firstNewObject + (request.pageContentConsolidation ? 1 : 0)
  ) {
    fail("precondition-mismatch", "Font object allocation changed.");
  }
  await assertFontPackageIntegrity(sourceBytes, parsed, mutation);

  const revisedTarget = revisedContentStreamObject(
    consolidatedBytes ?? sourceBytes,
    target,
    changedStream,
  );
  const { ref: resourcesRef, bytes: resourcesObject } =
    fontExtensionResourcesObject(graph, mutation);
  const pageSource = asLatin1(graph.pageObject.bytes);
  const pageWithResources = rewriteDictionaryEntry(
    pageSource,
    "Resources",
    `${resourcesRef.objectNumber} 0 R`,
    "PDF page Resources",
  );
  // A consolidated page also names its one new content stream.
  const revisedPageValue = latin1Bytes(
    request.pageContentConsolidation
      ? rewriteDictionaryEntry(
          pageWithResources,
          "Contents",
          `${request.pageContentConsolidation.stream.objectNumber} 0 R`,
          "PDF page Contents",
        )
      : pageWithResources,
  );
  const revisedPage =
    graph.pageObject.storage === "direct"
      ? revisedPageValue
      : concatBytes([
          ascii(
            `${graph.pageObject.objectNumber} ${graph.pageObject.generation} obj\n`,
          ),
          revisedPageValue,
          ascii("\nendobj"),
        ]);
  const revisions: IncrementalObjectRevision[] = [
    Object.freeze({
      objectNumber: request.identity.objectNumber,
      generation: request.identity.generation,
      bytes: revisedTarget,
    }),
    Object.freeze({
      objectNumber: mutation.page.objectNumber,
      generation: mutation.page.generation,
      bytes: revisedPage,
    }),
    ...pageFontPackageObjects(mutation).map((object) =>
      Object.freeze({
        objectNumber: object.objectNumber,
        generation: object.generation,
        bytes: object.bytes,
      }),
    ),
    Object.freeze({ ...resourcesRef, bytes: resourcesObject }),
  ].sort(
    (left, right) =>
      left.objectNumber - right.objectNumber ||
      left.generation - right.generation,
  );
  const incremental = makeMultiObjectIncrementalRevision(
    sourceBytes,
    parsed.current,
    revisions,
    {
      min: pageFontPackageObjects(mutation).length + 3,
      max: pageFontPackageObjects(mutation).length + 3,
      label: "Page Font extension",
    },
  );
  const preparedObject = async (revision: IncrementalObjectRevision) =>
    Object.freeze({
      objectNumber: revision.objectNumber,
      generation: revision.generation,
      offset: incremental.offsets.get(
        `${revision.objectNumber}:${revision.generation}`,
      )!,
      sha256: await sha256PdfBytes(revision.bytes),
    });
  const targetPrepared = await preparedObject(
    revisions.find(
      (revision) =>
        revision.objectNumber === request.identity.objectNumber &&
        revision.generation === request.identity.generation,
    )!,
  );
  return Object.freeze({
    candidateBytes: incremental.candidateBytes,
    changedObjectOffset: targetPrepared.offset,
    incrementalXrefOffset: incremental.incrementalXrefOffset,
    changedObjectSha256: targetPrepared.sha256,
    changedStreamSha256: await sha256PdfBytes(changedStream),
    changedDecodedStreamSha256: await sha256PdfBytes(changedDecodedStream),
    ...(request.pageContentConsolidation
      ? {
          pageContentConsolidation: Object.freeze({
            stream: targetPrepared,
            page: await preparedObject(
              revisions.find(
                (revision) =>
                  revision.objectNumber === mutation.page.objectNumber &&
                  revision.generation === mutation.page.generation,
              )!,
            ),
          }),
        }
      : {}),
    pageFontExtensionMutation: Object.freeze({
      resourceName: mutation.fontPackage.resourceName,
      font: mutation.fontPackage.font,
      fontObjects: Object.freeze(
        await Promise.all(
          pageFontPackageObjects(mutation).map((object) =>
            preparedObject(
              revisions.find(
                (revision) =>
                  revision.objectNumber === object.objectNumber &&
                  revision.generation === object.generation,
              )!,
            ),
          ),
        ),
      ),
      resources: await preparedObject(
        revisions.find(
          (revision) => revision.objectNumber === resourcesRef.objectNumber,
        )!,
      ),
      page: await preparedObject(
        revisions.find(
          (revision) =>
            revision.objectNumber === mutation.page.objectNumber &&
            revision.generation === mutation.page.generation,
        )!,
      ),
    }),
  });
}

/**
 * A write of the page's own text that draws in fonts the file holds and the
 * page did not name (PdfNativePageFontReferencePlan): three objects in one
 * revision -- the content stream, the page's Resources cloned as a new
 * object with the entries, and the page naming it. The plan the request
 * carries has to be the one the source gives now, to the digest.
 */
async function preparePageFontReference(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
  request: PdfNativeContentStreamTransactionRequest,
  target: ParsedIndirectObject,
  changedDecodedStream: Uint8Array,
  changedStream: Uint8Array,
): Promise<
  Readonly<{
    candidateBytes: Uint8Array;
    changedObjectOffset: number;
    incrementalXrefOffset: number;
    changedObjectSha256: string;
    changedStreamSha256: string;
    changedDecodedStreamSha256: string;
    pageFontReference: PdfNativePreparedPageFontReference;
  }>
> {
  const reference = request.pageFontReference;
  if (!reference) fail("invalid-request", "Page font reference is missing.");
  const resolved = await resolvePageFontReference(sourceBytes, parsed, reference.pageNumber, reference.fonts, request.identity);
  if (pageFontReferenceSummary(resolved.plan) !== pageFontReferenceSummary(reference)) {
    fail("precondition-mismatch", "Page Font reference authority changed.");
  }
  const revisions: IncrementalObjectRevision[] = [
    Object.freeze({
      objectNumber: request.identity.objectNumber,
      generation: request.identity.generation,
      bytes: revisedContentStreamObject(sourceBytes, target, changedStream),
    }),
    Object.freeze({
      objectNumber: reference.page.objectNumber,
      generation: reference.page.generation,
      bytes: resolved.revisedPage,
    }),
    Object.freeze({ ...resolved.plan.resources, bytes: resolved.resourcesObject }),
  ].sort((left, right) => left.objectNumber - right.objectNumber || left.generation - right.generation);
  const incremental = makeMultiObjectIncrementalRevision(sourceBytes, parsed.current, revisions, {
    min: 3,
    max: 3,
    label: "Page Font reference",
  });
  const preparedObject = async (objectNumber: number, generation: number) => {
    const revision = revisions.find((item) => item.objectNumber === objectNumber && item.generation === generation)!;
    return Object.freeze({
      objectNumber,
      generation,
      offset: incremental.offsets.get(`${objectNumber}:${generation}`)!,
      sha256: await sha256PdfBytes(revision.bytes),
    });
  };
  const targetPrepared = await preparedObject(request.identity.objectNumber, request.identity.generation);
  return Object.freeze({
    candidateBytes: incremental.candidateBytes,
    changedObjectOffset: targetPrepared.offset,
    incrementalXrefOffset: incremental.incrementalXrefOffset,
    changedObjectSha256: targetPrepared.sha256,
    changedStreamSha256: await sha256PdfBytes(changedStream),
    changedDecodedStreamSha256: await sha256PdfBytes(changedDecodedStream),
    pageFontReference: Object.freeze({
      fonts: Object.freeze(reference.fonts.map((item) => Object.freeze({
        resourceName: item.resourceName,
        font: Object.freeze({ objectNumber: item.font.objectNumber, generation: item.font.generation }),
      }))),
      resources: await preparedObject(resolved.plan.resources.objectNumber, 0),
      page: await preparedObject(reference.page.objectNumber, reference.page.generation),
    }),
  });
}

async function resolveContentAppendPage(
  sourceBytes: Uint8Array,
  identity: Readonly<{ objectNumber: number; generation: number }>,
  pageNumber: number,
): Promise<PdfNativePageContentGraphInspection> {
  const pages = await inspectMraiNativeDocumentPageGraphs(sourceBytes);
  const page = pages[pageNumber - 1];
  // A block is appended to the tail of the page drawing, and a page drawn
  // from several streams keeps that tail in its last one.
  if (
    !page ||
    page.pageNumber !== pageNumber ||
    page.contentStreams.length < 1 ||
    !sameNativeObjectRef(page.contentStreams.at(-1)!, identity) ||
    page.contentStreams.filter((stream) =>
      sameNativeObjectRef(stream, identity),
    ).length !== 1
  ) {
    fail(
      "unsupported-document",
      "Content append requires the page's own last content stream.",
    );
  }
  if (
    pages.some(
      (other) =>
        other !== page &&
        other.contentStreams.some((stream) =>
          sameNativeObjectRef(stream, identity),
        ),
    )
  ) {
    fail(
      "unsupported-document",
      "Content append stream is shared between pages.",
    );
  }
  return page;
}

function serializeContentAppendExtGState(
  reference: Readonly<{ objectNumber: number; generation: 0 }>,
  fillOpacity: number | null,
  strokeOpacity: number | null,
): Uint8Array {
  const entries = [
    fillOpacity === null ? null : `/ca ${formatNativePdfNumber(fillOpacity)}`,
    strokeOpacity === null
      ? null
      : `/CA ${formatNativePdfNumber(strokeOpacity)}`,
  ].filter((entry): entry is string => entry !== null);
  return ascii(
    `${reference.objectNumber} 0 obj\n<< /Type /ExtGState ${entries.join(" ")} >>\nendobj`,
  );
}

function contentAppendAdditionObjects(
  addition: PdfNativeContentAppendAddition,
): readonly IncrementalObjectRevision[] {
  if (addition.category === "Font") {
    return addition.fontPackage.objects.map((object) =>
      Object.freeze({
        objectNumber: object.objectNumber,
        generation: object.generation,
        bytes: object.bytes,
      }),
    );
  }
  if (addition.category === "XObject") {
    const normalized = normalizedEmbeddedImageReplacementData(addition.image);
    const image = serializeEmbeddedImageObject(
      addition.reference,
      normalized.evidence,
      normalized.encodedBytes,
      addition.softMask ?? undefined,
    );
    const mask =
      addition.softMask && normalized.softMaskBytes
        ? serializeEmbeddedImageObject(
            addition.softMask,
            {
              ...normalized.evidence,
              colorSpace: "DeviceGray",
              filter: "FlateDecode",
              decodeParms: {
                predictor: 15,
                colors: 1,
                bitsPerComponent: 8,
                columns: normalized.evidence.width,
              },
            },
            normalized.softMaskBytes,
          )
        : null;
    return [
      Object.freeze({ ...addition.reference, bytes: image }),
      ...(mask && addition.softMask
        ? [Object.freeze({ ...addition.softMask, bytes: mask })]
        : []),
    ];
  }
  return [
    Object.freeze({
      ...addition.reference,
      bytes: serializeContentAppendExtGState(
        addition.reference,
        addition.fillOpacity,
        addition.strokeOpacity,
      ),
    }),
  ];
}

async function prepareContentAppend(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
  request: PdfNativeContentStreamTransactionRequest,
  target: ParsedIndirectObject,
  changedDecodedStream: Uint8Array,
  changedStream: Uint8Array,
): Promise<
  Readonly<{
    candidateBytes: Uint8Array;
    changedObjectOffset: number;
    incrementalXrefOffset: number;
    changedObjectSha256: string;
    changedStreamSha256: string;
    changedDecodedStreamSha256: string;
    contentAppend: PdfNativePreparedContentAppend;
  }>
> {
  const plan = request.contentAppend;
  if (!plan) fail("invalid-request", "Content append plan is missing.");
  const page = await resolveContentAppendPage(
    sourceBytes,
    request.identity,
    plan.pageNumber,
  );
  if (
    plan.direction === "append" &&
    request.operatorSpan.start !== target.streamBytes!.length
  ) {
    fail(
      "invalid-request",
      "Content append span must sit at the end of the content stream.",
    );
  }
  const pageObject = parseIndirectObject(
    sourceBytes,
    parsed,
    page.page.objectNumber,
    page.page.generation,
  );
  const ownerObject = parseIndirectObject(
    sourceBytes,
    parsed,
    page.resourceAuthority.owner.objectNumber,
    page.resourceAuthority.owner.generation,
  );
  if (
    !sameNativeObjectRef(page.page, plan.page) ||
    (await sha256PdfBytes(pageObject.bytes)) !== plan.page.expectedObjectSha256 ||
    !sameNativeObjectRef(page.resourceAuthority.owner, plan.resourceOwner) ||
    (await sha256PdfBytes(ownerObject.bytes)) !==
      plan.resourceOwner.expectedObjectSha256 ||
    page.resourceAuthority.dictionaryKind !== plan.resourceOwner.dictionaryKind
  ) {
    fail(
      "precondition-mismatch",
      "Content append page resource authority changed.",
    );
  }
  const graphs = new Map<PdfNativeContentAppendCategory, ResolvedPageFontGraph>();
  const graphFor = async (category: PdfNativeContentAppendCategory) => {
    let graph = graphs.get(category);
    if (!graph) {
      graph = await resolvePageResourceCategoryGraph(
        sourceBytes,
        parsed,
        page,
        category,
      );
      graphs.set(category, graph);
    }
    return graph;
  };
  for (const category of CONTENT_APPEND_CATEGORIES) {
    const state = plan.categories[category];
    if (!state) continue;
    const graph = await graphFor(category);
    if (JSON.stringify(graph.category) !== JSON.stringify(state)) {
      fail(
        "precondition-mismatch",
        `Content append ${category} resources changed.`,
      );
    }
  }
  for (const addition of plan.additions) {
    const graph = await graphFor(addition.category);
    if (
      graph.entries.some(
        (entry) => entry.resourceName === addition.resourceName,
      ) ||
      graph.names.has(addition.resourceName)
    ) {
      fail(
        "precondition-mismatch",
        `Content append ${addition.category} resource name already exists.`,
      );
    }
  }
  for (const reused of plan.reusedResources) {
    const graph = await graphFor(reused.category);
    const entry = graph.entries.find(
      (candidate) => candidate.resourceName === reused.resourceName,
    );
    if (!entry || !sameNativeObjectRef(entry.target, reused.reference)) {
      fail(
        "precondition-mismatch",
        "Content append reused resource is missing or moved.",
      );
    }
    const object = parseIndirectObject(
      sourceBytes,
      parsed,
      reused.reference.objectNumber,
      reused.reference.generation,
    );
    if ((await sha256PdfBytes(object.bytes)) !== reused.expectedObjectSha256) {
      fail(
        "precondition-mismatch",
        "Content append reused resource object changed.",
      );
    }
  }
  const revisedTarget = revisedContentStreamObject(
    sourceBytes,
    target,
    changedStream,
  );
  if (plan.additions.length === 0) {
    const incremental = makeIncrementalRevision(
      sourceBytes,
      parsed.current,
      request.identity.objectNumber,
      request.identity.generation,
      revisedTarget,
    );
    return Object.freeze({
      candidateBytes: incremental.candidateBytes,
      changedObjectOffset: incremental.changedObjectOffset,
      incrementalXrefOffset: incremental.incrementalXrefOffset,
      changedObjectSha256: await sha256PdfBytes(revisedTarget),
      changedStreamSha256: await sha256PdfBytes(changedStream),
      changedDecodedStreamSha256: await sha256PdfBytes(changedDecodedStream),
      contentAppend: Object.freeze({
        objects: Object.freeze([]),
        resources: null,
        page: null,
      }),
    });
  }
  const latest = latestEntries(parsed);
  let cursor = Math.max(parsed.current.size, Math.max(0, ...latest.keys()) + 1);
  const additionRevisions: IncrementalObjectRevision[] = [];
  for (const addition of plan.additions) {
    for (const object of contentAppendAdditionObjects(addition)) {
      if (object.objectNumber !== cursor || object.generation !== 0) {
        fail(
          "precondition-mismatch",
          "Content append object allocation changed.",
        );
      }
      cursor += 1;
      additionRevisions.push(object);
    }
  }
  const resourcesRef = Object.freeze({ objectNumber: cursor, generation: 0 as const });
  if (resourcesRef.objectNumber > 9_999_999_999) {
    fail("unsupported-document", "Content append object number is too large.");
  }
  let resourcesDictionary = (await graphFor(plan.additions[0].category))
    .resourcesDictionary;
  for (const category of CONTENT_APPEND_CATEGORIES) {
    const additions = plan.additions.filter(
      (addition) => addition.category === category,
    );
    if (additions.length === 0) continue;
    const graph = await graphFor(category);
    let categoryDictionary = graph.categoryDictionary;
    for (const addition of additions) {
      const reference =
        addition.category === "Font"
          ? addition.fontPackage.font
          : addition.reference;
      categoryDictionary = rewriteDictionaryEntry(
        categoryDictionary,
        addition.resourceName,
        `${reference.objectNumber} 0 R`,
        `PDF ${category} resource name`,
      );
    }
    resourcesDictionary = rewriteDictionaryEntry(
      resourcesDictionary,
      category,
      categoryDictionary,
      `PDF ${category} category`,
    );
  }
  const resourcesObject = latin1Bytes(
    `${resourcesRef.objectNumber} 0 obj\n${resourcesDictionary}\nendobj`,
  );
  const revisedPageValue = latin1Bytes(
    rewriteDictionaryEntry(
      asLatin1(pageObject.bytes),
      "Resources",
      `${resourcesRef.objectNumber} 0 R`,
      "PDF page Resources",
    ),
  );
  const revisedPage =
    pageObject.storage === "direct"
      ? revisedPageValue
      : concatBytes([
          ascii(`${pageObject.objectNumber} ${pageObject.generation} obj\n`),
          revisedPageValue,
          ascii("\nendobj"),
        ]);
  const revisions: IncrementalObjectRevision[] = [
    Object.freeze({
      objectNumber: request.identity.objectNumber,
      generation: request.identity.generation,
      bytes: revisedTarget,
    }),
    Object.freeze({
      objectNumber: page.page.objectNumber,
      generation: page.page.generation,
      bytes: revisedPage,
    }),
    ...additionRevisions,
    Object.freeze({ ...resourcesRef, bytes: resourcesObject }),
  ].sort(
    (left, right) =>
      left.objectNumber - right.objectNumber ||
      left.generation - right.generation,
  );
  const incremental = makeMultiObjectIncrementalRevision(
    sourceBytes,
    parsed.current,
    revisions,
    // The stream, the page, its resources, and each addition's objects: a
    // line of several scripts installs a face of each (a package is at most
    // 16 objects; buildShapedTextBlock).
    { min: 4, max: 3 + 16 * 8, label: "Content append" },
  );
  const preparedObject = async (revision: IncrementalObjectRevision) =>
    Object.freeze({
      objectNumber: revision.objectNumber,
      generation: revision.generation,
      offset: incremental.offsets.get(
        `${revision.objectNumber}:${revision.generation}`,
      )!,
      sha256: await sha256PdfBytes(revision.bytes),
    });
  const find = (reference: Readonly<{ objectNumber: number; generation: number }>) =>
    revisions.find((revision) => sameNativeObjectRef(revision, reference))!;
  const targetPrepared = await preparedObject(find(request.identity));
  return Object.freeze({
    candidateBytes: incremental.candidateBytes,
    changedObjectOffset: targetPrepared.offset,
    incrementalXrefOffset: incremental.incrementalXrefOffset,
    changedObjectSha256: targetPrepared.sha256,
    changedStreamSha256: await sha256PdfBytes(changedStream),
    changedDecodedStreamSha256: await sha256PdfBytes(changedDecodedStream),
    contentAppend: Object.freeze({
      objects: Object.freeze(
        await Promise.all(
          additionRevisions.map((revision) => preparedObject(revision)),
        ),
      ),
      resources: await preparedObject(find(resourcesRef)),
      page: await preparedObject(find(page.page)),
    }),
  });
}

/**
 * The dictionaries the page's own resources name, by name, in one category:
 * a direct dictionary as written, an indirect one as its object writes it,
 * null for a name the category does not hold exactly once (or holds as
 * anything else). Read-only (mirror of native_text_group_move's
 * _page_resource_dictionaries).
 */
export async function readMraiNativePageResourceDictionaries(
  sourceBytes: Uint8Array,
  pageNumber: number,
  category: "Font" | "ExtGState",
  names: readonly string[],
): Promise<ReadonlyMap<string, string | null>> {
  const pages = await inspectMraiNativeDocumentPageGraphs(sourceBytes);
  const page = pages[pageNumber - 1];
  if (!page || page.pageNumber !== pageNumber) {
    fail("invalid-request", "Selected PDF page is out of range.");
  }
  const parsed = parsePdf(sourceBytes, false);
  const graph = await resolvePageResourceCategoryGraph(sourceBytes, parsed, page, category);
  const entries =
    graph.categoryDictionary === "<<>>" ? [] : topLevelDictionaryEntries(graph.categoryDictionary);
  const found = new Map<string, string | null>();
  for (const name of names) {
    const matches = entries.filter((entry) => entry.key === name);
    let dictionary: string | null = null;
    if (matches.length === 1) {
      const entry = matches[0];
      const value = graph.categoryDictionary.slice(entry.valueStart, entry.valueEnd);
      try {
        if (value.startsWith("<<")) {
          dictionary = value;
        } else {
          const reference = parseTopLevelReference(
            graph.categoryDictionary, entry.valueStart, entry.valueEnd, `PDF page ${category} resource`,
          );
          const object = parseIndirectObject(
            sourceBytes, parsed, reference.reference.objectNumber, reference.reference.generation,
          );
          dictionary = reference.end === entry.valueEnd && object.streamBytes === null
            ? topLevelDictionarySource(asLatin1(object.bytes), `PDF page ${category} resource`)
            : null;
        }
      } catch (error) {
        if (!(error instanceof PdfNativeTransactionError)) throw error;
        dictionary = null;
      }
    }
    found.set(name, dictionary);
  }
  return found;
}

/**
 * Read-only page authority for planning one content append: the single
 * exclusively owned content stream, the effective Resources owner and the
 * current state of every resource category the block may extend.
 */
export async function inspectMraiNativeContentAppendPage(
  sourceBytes: Uint8Array,
  pageNumber: number,
): Promise<PdfNativeContentAppendPageInspection> {
  const exactPage = exactPositiveInteger(pageNumber, "content append page");
  const parsed = parsePdf(sourceBytes, false);
  const sourceSha256 = await sha256PdfBytes(sourceBytes);
  const pages = await inspectMraiNativeDocumentPageGraphs(sourceBytes);
  const page = pages[exactPage - 1];
  if (!page || page.pageNumber !== exactPage) {
    fail("invalid-request", "Selected PDF page is out of range.");
  }
  if (page.contentStreams.length < 1) {
    fail(
      "unsupported-document",
      "Content append requires the page's own last content stream.",
    );
  }
  const stream = page.contentStreams.at(-1)!;
  await resolveContentAppendPage(sourceBytes, stream, exactPage);
  const inspections = await inspectMraiNativePageContentStreams(
    sourceBytes,
    exactPage,
  );
  const target = inspections.at(-1)!;
  const pageObject = parseIndirectObject(
    sourceBytes,
    parsed,
    page.page.objectNumber,
    page.page.generation,
  );
  const ownerObject = parseIndirectObject(
    sourceBytes,
    parsed,
    page.resourceAuthority.owner.objectNumber,
    page.resourceAuthority.owner.generation,
  );
  const categories = {} as Record<
    PdfNativeContentAppendCategory,
    PdfNativeContentAppendCategoryInspection
  >;
  for (const category of CONTENT_APPEND_CATEGORIES) {
    const graph = await resolvePageResourceCategoryGraph(
      sourceBytes,
      parsed,
      page,
      category,
    );
    categories[category] = Object.freeze({
      state: graph.category,
      entries: graph.entries,
    });
  }
  const latest = latestEntries(parsed);
  return Object.freeze({
    sourceSha256,
    pageNumber: exactPage,
    page: Object.freeze({
      ...page.page,
      expectedObjectSha256: await sha256PdfBytes(pageObject.bytes),
    }),
    resourceOwner: Object.freeze({
      ...page.resourceAuthority.owner,
      expectedObjectSha256: await sha256PdfBytes(ownerObject.bytes),
      dictionaryKind: page.resourceAuthority.dictionaryKind,
    }),
    contentStream: target,
    drawingPrefix: joinNativePageDrawing(inspections.slice(0, -1)),
    categories: Object.freeze(categories),
    firstNewObjectNumber: Math.max(
      parsed.current.size,
      Math.max(0, ...latest.keys()) + 1,
    ),
  });
}

/** The refusal for an edit that would change the glyphs an /ActualText
 * cluster span stands over (see actualTextClustersKeepTheirShows). */
export const MRAI_NATIVE_ACTUAL_TEXT_CLUSTER_MESSAGE =
  "The edit would change glyphs an /ActualText names; readers that copy by it would read the old letters.";

/** The bytes of a PDF literal string's value from its opening parenthesis,
 * escapes resolved (ISO 32000-1 7.3.4.2); null when it does not close. */
function pdfLiteralStringValue(source: string, start: number): Readonly<{ bytes: number[]; end: number }> | null {
  const bytes: number[] = [];
  let depth = 0;
  for (let cursor = start; cursor < source.length; cursor += 1) {
    const character = source[cursor];
    if (character === "\\") {
      const escaped = source[cursor + 1];
      if (escaped === undefined) return null;
      const named = ({ n: 0x0a, r: 0x0d, t: 0x09, b: 0x08, f: 0x0c } as Record<string, number>)[escaped];
      if (named !== undefined) { bytes.push(named); cursor += 1; continue; }
      if (/[0-7]/u.test(escaped)) {
        let octal = escaped;
        while (octal.length < 3 && /[0-7]/u.test(source[cursor + 1 + octal.length] ?? "")) octal += source[cursor + 1 + octal.length];
        bytes.push(Number.parseInt(octal, 8) & 0xff);
        cursor += octal.length;
        continue;
      }
      if (escaped === "\n") { cursor += 1; continue; }
      if (escaped === "\r") { cursor += source[cursor + 2] === "\n" ? 2 : 1; continue; }
      bytes.push(escaped.charCodeAt(0));
      cursor += 1;
      continue;
    }
    if (character === "(") {
      depth += 1;
      if (depth === 1) continue;
    } else if (character === ")") {
      depth -= 1;
      if (depth === 0) return { bytes, end: cursor + 1 };
    }
    bytes.push(character.charCodeAt(0));
  }
  return null;
}

/**
 * Whether a marked-content property list's /ActualText names a cluster of
 * characters: two or more, none of them white space, as a literal or hex
 * string (UTF-16BE after its byte order mark, or printable ASCII). Chromium
 * wraps each ligature glyph in `/Span << /ActualText (fi) >> BDC`. The
 * reader reads text by the glyphs a show draws, never by an /ActualText
 * (pdf.js getTextContent), so such a glyph reads as its font names it ("fi"
 * from U+FB01) and is text like any other; the span only speaks to readers
 * that copy by /ActualText, and an edit never changes what it draws
 * (actualTextClustersKeepTheirShows). A span naming one character or white
 * space -- Word's tab or space glyph, a form's underscore, an optional
 * hyphen, Qt's empty overprint span, a leader read as ". . . ." -- keeps its
 * own rules (but see pdfNativeActualTextReadsAsText). Mirrors the server's
 * _actual_text_cluster_value.
 */
export function pdfNativeActualTextClusterValue(properties: string): boolean {
  const characters = pdfNativeActualTextCharacters(properties);
  return characters !== null && characters.length >= 2 && characters.every((character) => /\S/u.test(character));
}

/** The characters a marked-content property list's /ActualText names (a
 * literal or hex string: UTF-16BE after its byte order mark, or printable
 * ASCII); null when there is none this reads. */
function pdfNativeActualTextCharacters(properties: string): string[] | null {
  const key = /\/ActualText\s*/u.exec(properties);
  if (!key) return null;
  const start = key.index + key[0].length;
  let bytes: number[];
  if (properties[start] === "(") {
    const literal = pdfLiteralStringValue(properties, start);
    if (!literal) return null;
    bytes = literal.bytes;
  } else if (properties[start] === "<" && properties[start + 1] !== "<") {
    const end = properties.indexOf(">", start);
    const hex = end < 0 ? "" : properties.slice(start + 1, end).replace(/\s+/gu, "");
    if (end < 0 || !/^[0-9A-Fa-f]*$/u.test(hex)) return null;
    const even = hex.length % 2 ? `${hex}0` : hex;
    bytes = (even.match(/../gu) ?? []).map((pair) => Number.parseInt(pair, 16));
  } else {
    return null;
  }
  let text: string;
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
    if (bytes.length % 2) return null;
    text = "";
    for (let index = 2; index < bytes.length; index += 2) text += String.fromCharCode((bytes[index] << 8) | bytes[index + 1]);
  } else {
    if (!bytes.every((byte) => byte >= 0x21 && byte <= 0x7e)) return null;
    text = String.fromCharCode(...bytes);
  }
  return Array.from(text);
}

/**
 * Whether the reader reads the glyphs under an /ActualText as text: a
 * cluster's (pdfNativeActualTextClusterValue), or one ideograph's that a
 * radical glyph stands for (pdfNativeRadicalStandsFor). Only a cluster's
 * glyphs are held to what they draw (actualTextClustersKeepTheirShows); a
 * radical span is left to the rules a text box's hiding keeps, as before.
 */
export function pdfNativeActualTextReadsAsText(properties: string): boolean {
  if (pdfNativeActualTextClusterValue(properties)) return true;
  const characters = pdfNativeActualTextCharacters(properties);
  return characters !== null && characters.length === 1 && pdfNativeRadicalStandsFor(characters[0]);
}

let radicalTargets: ReadonlySet<string> | null = null;










export function pdfNativeRadicalStandsFor(character: string): boolean {
  if (!radicalTargets) {
    const targets = new Set<string>();
    const ranges: ReadonlyArray<readonly [number, number]> = [[0x2e80, 0x2ef3], [0x2f00, 0x2fd5], [0xf900, 0xfaff], [0x2f800, 0x2fa1f]];
    for (const [first, last] of ranges) {
      for (let code = first; code <= last; code += 1) {
        const form = String.fromCodePoint(code).normalize("NFKC");
        if (form !== String.fromCodePoint(code) && Array.from(form).length === 1 && /\p{Script=Han}/u.test(form)) targets.add(form);
      }
    }
    radicalTargets = targets;
  }
  return radicalTargets.has(character);
}

const CONTENT_WHITE_SPACE = new Set(["\0", "\t", "\n", "\f", "\r", " "]);
const CONTENT_DELIMITERS = "()<>[]{}/%";

/** An /ActualText cluster span of a content stream: its opening (the
 * `BDC` with its operands), its closing `EMC`, and each show under it with
 * the codes it draws (its strings' bytes, in hex). */
export interface PdfNativeActualTextClusterSpan {
  readonly openStart: number;
  readonly openEnd: number;
  readonly closeStart: number;
  readonly closeEnd: number;
  readonly shows: readonly Readonly<{ start: number; end: number; drawn: string }>[];
}

/**
 * Every /ActualText cluster span of a content stream, the outermost, in
 * stream order. Null when this lexer cannot walk the stream to its end. A
 * span opened before the stretch it is given, or closed after it, is none of
 * its spans.
 */
export function pdfNativeActualTextClusterSpans(source: string): PdfNativeActualTextClusterSpan[] | null {
  const spans: PdfNativeActualTextClusterSpan[] = [];
  const stack: { cluster: boolean; openStart: number; openEnd: number; shows: { start: number; end: number; drawn: string }[] }[] = [];
  let operandStart = -1;
  let strings: string[] = [];
  let cursor = 0;
  const operand = (start: number) => { if (operandStart < 0) operandStart = start; };
  const hexOf = (bytes: readonly number[]) => bytes.map((byte) => byte.toString(16).padStart(2, "0")).join("");
  while (cursor < source.length) {
    const character = source[cursor];
    if (CONTENT_WHITE_SPACE.has(character)) { cursor += 1; continue; }
    if (character === "%") {
      while (cursor < source.length && source[cursor] !== "\n" && source[cursor] !== "\r") cursor += 1;
      continue;
    }
    const start = cursor;
    if (character === "(") {
      const literal = pdfLiteralStringValue(source, cursor);
      if (!literal) return null;
      cursor = literal.end;
      strings.push(hexOf(literal.bytes));
      operand(start);
      continue;
    }
    if (character === "<" && source[cursor + 1] === "<") {
      let depth = 0;
      while (cursor < source.length) {
        if (source[cursor] === "(") {
          const literal = pdfLiteralStringValue(source, cursor);
          if (!literal) return null;
          cursor = literal.end;
        } else if (source.startsWith("<<", cursor)) { depth += 1; cursor += 2; }
        else if (source.startsWith(">>", cursor)) { depth -= 1; cursor += 2; if (depth === 0) break; }
        else if (source[cursor] === "<") {
          const end = source.indexOf(">", cursor);
          if (end < 0) return null;
          cursor = end + 1;
        } else cursor += 1;
      }
      if (depth !== 0) return null;
      operand(start);
      continue;
    }
    if (character === "<") {
      const end = source.indexOf(">", cursor);
      if (end < 0) return null;
      const hex = source.slice(start + 1, end).replace(/\s+/gu, "").toLowerCase();
      strings.push(hex.length % 2 ? `${hex}0` : hex);
      cursor = end + 1;
      operand(start);
      continue;
    }
    if (character === "[") { cursor += 1; operand(start); continue; }
    if (character === "]" || character === "{" || character === "}" || character === ">" || character === ")") {
      cursor += 1;
      continue;
    }
    if (character === "/") {
      cursor += 1;
      while (cursor < source.length && !CONTENT_WHITE_SPACE.has(source[cursor]) && !CONTENT_DELIMITERS.includes(source[cursor])) cursor += 1;
      operand(start);
      continue;
    }
    while (cursor < source.length && !CONTENT_WHITE_SPACE.has(source[cursor]) && !CONTENT_DELIMITERS.includes(source[cursor])) cursor += 1;
    const word = source.slice(start, cursor);
    if (/^[-+.\d]/u.test(word) || word === "true" || word === "false" || word === "null") { operand(start); continue; }
    const operatorStart = operandStart < 0 ? start : operandStart;
    const drawn = strings.join("");
    operandStart = -1;
    strings = [];
    if (word === "BI") {
      // An inline image's data runs from `ID` and one white-space byte to an
      // `EI` standing alone.
      const data = /\sID\s/u.exec(source.slice(cursor));
      if (!data) return null;
      const from = cursor + data.index + data[0].length;
      const end = /\sEI(?=[\s]|$)/u.exec(source.slice(from));
      if (!end) return null;
      cursor = from + end.index + end[0].length;
      continue;
    }
    if (word === "BDC") {
      const opening = source.slice(operatorStart, cursor);
      stack.push({ cluster: /\/ActualText\b/u.test(opening) && pdfNativeActualTextClusterValue(opening),
        openStart: operatorStart, openEnd: cursor, shows: [] });
    } else if (word === "BMC") {
      stack.push({ cluster: false, openStart: operatorStart, openEnd: cursor, shows: [] });
    } else if (word === "EMC") {
      const closed = stack.pop();
      if (closed?.cluster && !stack.some((entry) => entry.cluster)) {
        spans.push(Object.freeze({ openStart: closed.openStart, openEnd: closed.openEnd, closeStart: start, closeEnd: cursor,
          shows: Object.freeze(closed.shows.map((show) => Object.freeze(show))) }));
      }
    } else if (word === "Tj" || word === "TJ" || word === "'" || word === "\"") {
      stack.find((entry) => entry.cluster)?.shows.push({ start: operatorStart, end: cursor, drawn });
    }
  }
  return spans;
}

/** Every /ActualText cluster span of a content stream as its opening and
 * the codes its shows draw: where a show stands, its spacing and the way its
 * operands are written are not what the span names. */
function actualTextClusterShows(source: string, before = Number.POSITIVE_INFINITY): string[] | null {
  const spans = pdfNativeActualTextClusterSpans(source);
  return spans && spans.filter((span) => span.openStart < before).map((span) =>
    `${source.slice(span.openStart, span.openEnd)}\u0000${span.shows.map((show) => show.drawn).join("\u0001")}`);
}

/**
 * Whether an edited content stream keeps every /ActualText cluster span of
 * its source whole: each span it still has is one of the source's, its
 * opening byte for byte and its shows drawing the same codes, in the same
 * order. An edit may move such a span's glyphs (a `Td` inside it is the
 * pen's), respace or rewrite their operators, or drop the span whole, but
 * never change what it draws, which would leave its /ActualText naming
 * letters the page no longer shows.
 *
 * `appendedFrom`: the edit appends a block from this offset on (a content
 * append, the stream's end): a span the block opens names its own new
 * glyphs (a shaped complex-script run's words), not the source's.
 */
export function actualTextClustersKeepTheirShows(original: Uint8Array, edited: Uint8Array,
  appendedFrom?: number): boolean {
  const source = LATIN1.decode(original);
  if (!source.includes("/ActualText")) return true;
  const before = actualTextClusterShows(source);
  // A source this lexer cannot walk is judged by the routes that proved it.
  if (!before || before.length === 0) return true;
  const after = actualTextClusterShows(LATIN1.decode(edited), appendedFrom);
  if (!after) return false;
  let cursor = 0;
  for (const span of after) {
    while (cursor < before.length && before[cursor] !== span) cursor += 1;
    if (cursor === before.length) return false;
    cursor += 1;
  }
  return true;
}

// ---------------------------------------------------------------------------
// A structure element's /ActualText, written with the content it names.
// ---------------------------------------------------------------------------













export interface PdfNativeStructActualTextRevision {
  readonly objectNumber: number;
  readonly generation: number;
  /** The element as the source holds it (its value alone when it sits in
   * an object stream), the way every object precondition is digested. */
  readonly expectedObjectSha256: string;
  /** The page the element names for its marked content. */
  readonly page: Readonly<{ objectNumber: number; generation: number }>;
  /** Every MCID the element's /K names, ascending. */
  readonly mcids: readonly number[];
  /** The /ActualText value token as the element holds it now. */
  readonly previous: Uint8Array;
  /** The token written in its place. */
  readonly value: Uint8Array;
  /** What `value` reads as. */
  readonly text: string;
}

const MAX_STRUCT_ACTUAL_TEXT_REVISIONS = 64;
const MAX_STRUCT_ACTUAL_TEXT_TOKEN_BYTES = 32 * 1024;

/** PDFDocEncoding's code points that are not Latin-1's (ISO 32000-1
 * Table D.2); 0x7F, 0x9F and 0xAD name no character, nor do the controls
 * other than tab, line feed and carriage return. Mirror of the server's
 * _PDF_DOC_ENCODING_HIGH. */
const PDF_DOC_ENCODING_HIGH: ReadonlyMap<number, number> = new Map([
  [0x18, 0x02d8], [0x19, 0x02c7], [0x1a, 0x02c6], [0x1b, 0x02d9],
  [0x1c, 0x02dd], [0x1d, 0x02db], [0x1e, 0x02da], [0x1f, 0x02dc],
  [0x80, 0x2022], [0x81, 0x2020], [0x82, 0x2021], [0x83, 0x2026],
  [0x84, 0x2014], [0x85, 0x2013], [0x86, 0x0192], [0x87, 0x2044],
  [0x88, 0x2039], [0x89, 0x203a], [0x8a, 0x2212], [0x8b, 0x2030],
  [0x8c, 0x201e], [0x8d, 0x201c], [0x8e, 0x201d], [0x8f, 0x2018],
  [0x90, 0x2019], [0x91, 0x201a], [0x92, 0x2122], [0x93, 0xfb01],
  [0x94, 0xfb02], [0x95, 0x0141], [0x96, 0x0152], [0x97, 0x0160],
  [0x98, 0x0178], [0x99, 0x017d], [0x9a, 0x0131], [0x9b, 0x0142],
  [0x9c, 0x0153], [0x9d, 0x0161], [0x9e, 0x017e], [0xa0, 0x20ac],
]);
const PDF_DOC_ENCODING_CODES: ReadonlyMap<number, number> = new Map(
  [...PDF_DOC_ENCODING_HIGH].map(([byte, code]) => [code, byte]),
);

function pdfDocEncodingCodePoint(byte: number): number | null {
  const high = PDF_DOC_ENCODING_HIGH.get(byte);
  if (high !== undefined) return high;
  if (byte === 0x09 || byte === 0x0a || byte === 0x0d) return byte;
  if ((byte >= 0x20 && byte <= 0x7e) || (byte >= 0xa1 && byte <= 0xff && byte !== 0xad)) return byte;
  return null;
}

/** The bytes a PDF string token (`(…)` or `<…>`) holds, or null when the
 * token is not one string (mirror of the server's _pdf_string_token_bytes). */
export function decodeMraiNativePdfStringToken(token: Uint8Array): Uint8Array | null {
  const source = asLatin1(token);
  if (source.startsWith("<") && !source.startsWith("<<")) {
    if (!source.endsWith(">")) return null;
    const digits = source.slice(1, -1).replace(/[\x00\x09\x0a\x0c\x0d\x20]/g, "");
    if (!/^[0-9A-Fa-f]*$/.test(digits)) return null;
    const padded = digits.length % 2 ? `${digits}0` : digits;
    const out = new Uint8Array(padded.length / 2);
    for (let index = 0; index < out.length; index += 1) {
      out[index] = Number.parseInt(padded.slice(index * 2, index * 2 + 2), 16);
    }
    return out;
  }
  if (!source.startsWith("(")) return null;
  const out: number[] = [];
  let depth = 0;
  for (let cursor = 0; cursor < source.length; cursor += 1) {
    const character = source[cursor];
    const code = source.charCodeAt(cursor);
    if (character === "\\") {
      const next = source[cursor + 1];
      if (next === undefined) return null;
      const named: Record<string, number> = { n: 0x0a, r: 0x0d, t: 0x09, b: 0x08, f: 0x0c };
      if (named[next] !== undefined) { out.push(named[next]); cursor += 1; continue; }
      if (next === "\r") { cursor += source[cursor + 2] === "\n" ? 2 : 1; continue; }
      if (next === "\n") { cursor += 1; continue; }
      if (/[0-7]/.test(next)) {
        let octal = "";
        while (octal.length < 3 && /[0-7]/.test(source[cursor + 1] ?? "")) {
          octal += source[cursor + 1];
          cursor += 1;
        }
        out.push(Number.parseInt(octal, 8) & 0xff);
        continue;
      }
      out.push(next.charCodeAt(0));
      cursor += 1;
      continue;
    }
    if (character === "(") {
      depth += 1;
      if (depth > 1) out.push(code);
      continue;
    }
    if (character === ")") {
      depth -= 1;
      if (depth === 0) return cursor === source.length - 1 ? Uint8Array.from(out) : null;
      out.push(code);
      continue;
    }
    // An end of line inside a literal reads as one line feed.
    if (character === "\r") {
      out.push(0x0a);
      if (source[cursor + 1] === "\n") cursor += 1;
      continue;
    }
    out.push(code);
  }
  return null;
}

/** A PDF text string's characters (ISO 32000-1 §7.9.2.2): UTF-16BE after
 * its byte order mark, else PDFDocEncoding. `strict` refuses a byte that
 * names no character and a broken UTF-16 sequence (null); otherwise such a
 * byte reads as the character Latin-1 gives it. Mirror of the server's
 * _pdf_text_string. */
export function decodeMraiNativePdfTextString(bytes: Uint8Array, strict = true): string | null {
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
    if (bytes.length % 2) return null;
    let text = "";
    for (let index = 2; index < bytes.length; index += 2) {
      text += String.fromCharCode((bytes[index] << 8) | bytes[index + 1]);
    }
    if (strict && !isWellFormedUtf16(text)) return null;
    return text;
  }
  let text = "";
  for (const byte of bytes) {
    const code = pdfDocEncodingCodePoint(byte);
    if (code === null) {
      if (strict) return null;
      text += String.fromCharCode(byte);
    } else {
      text += String.fromCharCode(code);
    }
  }
  return text;
}

function isWellFormedUtf16(text: string): boolean {
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index);
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = text.charCodeAt(index + 1);
      if (!(next >= 0xdc00 && next <= 0xdfff)) return false;
      index += 1;
    } else if (code >= 0xdc00 && code <= 0xdfff) {
      return false;
    }
  }
  return true;
}

/**
 * `text` as a PDF text string token written the way `previous` was: a
 * literal stays a literal and a hex string a hex string (in its case);
 * UTF-16BE with a byte order mark when the previous string was, or when the
 * text holds a character PDFDocEncoding cannot spell, else PDFDocEncoding.
 * Null when the text cannot be written (a broken UTF-16 sequence).
 */
export function encodeMraiNativePdfTextStringToken(text: string, previous: Uint8Array): Uint8Array | null {
  if (!isWellFormedUtf16(text)) return null;
  const previousBytes = decodeMraiNativePdfStringToken(previous);
  if (!previousBytes) return null;
  const previousSource = asLatin1(previous);
  const hex = previousSource.startsWith("<");
  const unicode = previousBytes.length >= 2 && previousBytes[0] === 0xfe && previousBytes[1] === 0xff;
  let bytes: number[] | null = unicode ? null : [];
  if (bytes) {
    for (const character of text) {
      const code = character.codePointAt(0)!;
      const mapped = PDF_DOC_ENCODING_CODES.get(code) ??
        (code <= 0xff && pdfDocEncodingCodePoint(code) === code ? code : undefined);
      if (mapped === undefined) {
        bytes = null;
        break;
      }
      bytes.push(mapped);
    }
  }
  if (!bytes) {
    bytes = [0xfe, 0xff];
    for (let index = 0; index < text.length; index += 1) {
      const code = text.charCodeAt(index);
      bytes.push(code >> 8, code & 0xff);
    }
  }
  if (hex) {
    const upper = !/[a-f]/.test(previousSource);
    const digits = bytes.map((byte) => byte.toString(16).padStart(2, "0")).join("");
    return latin1Bytes(`<${upper ? digits.toUpperCase() : digits}>`);
  }
  let literal = "(";
  for (const byte of bytes) {
    if (byte === 0x28 || byte === 0x29 || byte === 0x5c) literal += `\\${String.fromCharCode(byte)}`;
    else if (byte === 0x0a) literal += "\\n";
    else if (byte === 0x0d) literal += "\\r";
    else if (byte === 0x09) literal += "\\t";
    else if (byte >= 0x20 && byte <= 0x7e) literal += String.fromCharCode(byte);
    else literal += `\\${byte.toString(8).padStart(3, "0")}`;
  }
  return latin1Bytes(`${literal})`);
}

/** Where a structure element names its /ActualText: the value token's span
 * in the object's bytes, and the token. Null when the object is not one
 * dictionary naming exactly one /ActualText string. */
export function readMraiNativeStructElementActualText(
  objectBytes: Uint8Array,
): Readonly<{ start: number; end: number; token: Uint8Array }> | null {
  const source = asLatin1(objectBytes);
  let entries: readonly TopLevelDictionaryEntry[];
  try {
    entries = topLevelDictionaryEntries(source);
  } catch (error) {
    if (error instanceof PdfNativeTransactionError) return null;
    throw error;
  }
  const matches = entries.filter((entry) => entry.key === "ActualText");
  if (matches.length !== 1) return null;
  const { valueStart, valueEnd } = matches[0];
  const token = objectBytes.slice(valueStart, valueEnd);
  if (!decodeMraiNativePdfStringToken(token)) return null;
  return Object.freeze({ start: valueStart, end: valueEnd, token });
}

/** The element written again with `value` in place of its /ActualText value
 * and nothing else changed; a member of an object stream is written under a
 * header of its own (mirror of the server's _struct_actual_text_object). */
function structActualTextRevisedObject(
  object: ParsedIndirectObject,
  revision: PdfNativeStructActualTextRevision,
): Uint8Array {
  if (object.streamStart !== null) {
    fail("unsupported-document", "A structure element is a dictionary, not a stream.");
  }
  const found = readMraiNativeStructElementActualText(object.bytes);
  if (!found || !bytesEqual(found.token, revision.previous)) {
    fail("precondition-mismatch", "The structure element's /ActualText changed.");
  }
  const revised = concatBytes([
    object.bytes.subarray(0, found.start),
    revision.value,
    object.bytes.subarray(found.end),
  ]);
  return object.storage === "direct"
    ? revised
    : concatBytes([
        ascii(`${object.objectNumber} ${object.generation} obj\n`),
        revised,
        ascii("\nendobj"),
      ]);
}

function normalizeStructActualText(
  request: PdfNativeContentStreamTransactionRequest,
): readonly PdfNativeStructActualTextRevision[] | null {
  const revisions = request.structActualText;
  if (revisions === undefined) return null;
  if (!Array.isArray(revisions) || revisions.length < 1 || revisions.length > MAX_STRUCT_ACTUAL_TEXT_REVISIONS) {
    fail("invalid-request", "Structure /ActualText revisions are invalid.");
  }
  if (request.formInstanceIsolation || request.contentAppend) {
    fail("invalid-request", "A structure /ActualText rides a page's own text write.");
  }
  const seen = new Set<string>([`${request.identity.objectNumber}:${request.identity.generation}`]);
  for (const item of request.untouchedObjects ?? []) seen.add(`${item.objectNumber}:${item.generation}`);
  const ordered = [...revisions].sort((left, right) =>
    left.objectNumber - right.objectNumber || left.generation - right.generation);
  return Object.freeze(ordered.map((revision) => {
    assertObjectReference(revision.objectNumber, revision.generation);
    assertObjectReference(revision.page.objectNumber, revision.page.generation);
    const key = `${revision.objectNumber}:${revision.generation}`;
    if (seen.has(key)) fail("invalid-request", "A structure element is revised twice or is pinned untouched.");
    seen.add(key);
    const mcids = [...revision.mcids];
    if (
      mcids.length < 1 || mcids.length > 4_096 ||
      mcids.some((mcid, index) => !Number.isSafeInteger(mcid) || mcid < 0 || mcid > 1_000_000_000 ||
        (index > 0 && mcid <= mcids[index - 1])) ||
      !isUint8Array(revision.previous) || !isUint8Array(revision.value) ||
      revision.previous.length > MAX_STRUCT_ACTUAL_TEXT_TOKEN_BYTES ||
      revision.value.length > MAX_STRUCT_ACTUAL_TEXT_TOKEN_BYTES ||
      typeof revision.text !== "string"
    ) {
      fail("invalid-request", "A structure /ActualText revision is invalid.");
    }
    const written = decodeMraiNativePdfStringToken(revision.value);
    // Read as a reader reads it: a history writes back the token the
    // element held, which may name a byte PDFDocEncoding leaves undefined.
    if (!written || decodeMraiNativePdfTextString(written, false) !== revision.text || !decodeMraiNativePdfStringToken(revision.previous)) {
      fail("invalid-request", "A structure /ActualText revision does not read as its text.");
    }
    return Object.freeze({
      objectNumber: revision.objectNumber,
      generation: revision.generation,
      expectedObjectSha256: exactSha256(revision.expectedObjectSha256, "structure element"),
      page: Object.freeze({ objectNumber: revision.page.objectNumber, generation: revision.page.generation }),
      mcids: Object.freeze(mcids),
      previous: Uint8Array.from(revision.previous),
      value: Uint8Array.from(revision.value),
      text: revision.text,
    });
  }));
}

/**
 * The candidate with each structure element's revision appended after every
 * object the content write revised, and one cross-reference section listing
 * those objects and these. The rows the content write appended are read
 * back from its own section, so every writer's revision -- one stream, a
 * Form's clones, an installed font, a joined page -- takes them the same
 * way (mirror of the server's _append_struct_actual_text).
 */
async function appendStructActualTextRevisions(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
  base: Readonly<{ candidateBytes: Uint8Array; incrementalXrefOffset: number }>,
  revisions: readonly PdfNativeStructActualTextRevision[],
): Promise<Readonly<{
  candidateBytes: Uint8Array;
  incrementalXrefOffset: number;
  objects: readonly PdfNativePreparedIsolatedObject[];
}>> {
  const rows = appendedRevisionRows(parsed, base);
  const written: { objectNumber: number; generation: number; bytes: Uint8Array }[] = [];
  for (const revision of revisions) {
    if (rows.some((row) => row.objectNumber === revision.objectNumber)) {
      fail("invalid-request", "A structure element is an object the content write revises.");
    }
    const object = parseIndirectObject(sourceBytes, parsed, revision.objectNumber, revision.generation, undefined, "opaque");
    if ((await sha256PdfBytes(object.bytes)) !== revision.expectedObjectSha256) {
      fail("precondition-mismatch", "The structure element changed.");
    }
    written.push({
      objectNumber: revision.objectNumber,
      generation: revision.generation,
      bytes: structActualTextRevisedObject(object, revision),
    });
  }
  return appendObjectsToRevision(parsed, base, rows, written);
}

/**
 * The rows a candidate's one appended revision lists (every object it wrote,
 * read back from its own section), checked to be that revision over the
 * source's newest section.
 */
function appendedRevisionRows(
  parsed: ParsedPdf,
  base: Readonly<{ candidateBytes: Uint8Array; incrementalXrefOffset: number }>,
): { objectNumber: number; generation: number; offset: number }[] {
  const written = parsePdf(base.candidateBytes, true);
  if (written.current.offset !== base.incrementalXrefOffset || written.previous?.offset !== parsed.current.offset) {
    fail("verification-failed", "The content write's revision cannot be read back.");
  }
  const rows: { objectNumber: number; generation: number; offset: number }[] = [];
  for (const entry of written.current.entries.values()) {
    if (!entry.inUse || entry.kind !== "direct") continue;
    // A cross-reference stream lists itself; the section written below
    // replaces it.
    if (entry.offset === base.incrementalXrefOffset) continue;
    rows.push({ objectNumber: entry.objectNumber, generation: entry.generation, offset: entry.offset });
  }
  return rows;
}

/**
 * `base` with `objects` written after the objects its revision wrote
 * (`rows`), and one cross-reference section listing those and these (mirror
 * of the server's append_revision_objects).
 */
async function appendObjectsToRevision(
  parsed: ParsedPdf,
  base: Readonly<{ candidateBytes: Uint8Array; incrementalXrefOffset: number }>,
  rows: { objectNumber: number; generation: number; offset: number }[],
  objects: readonly Readonly<{ objectNumber: number; generation: number; bytes: Uint8Array }>[],
): Promise<Readonly<{
  candidateBytes: Uint8Array;
  incrementalXrefOffset: number;
  objects: readonly PdfNativePreparedIsolatedObject[];
}>> {
  const parts: Uint8Array[] = [base.candidateBytes.subarray(0, base.incrementalXrefOffset)];
  let cursor = base.incrementalXrefOffset;
  const prepared: PdfNativePreparedIsolatedObject[] = [];
  for (const object of objects) {
    if (cursor > 9_999_999_999) fail("unsupported-document", "Incremental object offset is too large.");
    rows.push({ objectNumber: object.objectNumber, generation: object.generation, offset: cursor });
    prepared.push(Object.freeze({
      objectNumber: object.objectNumber,
      generation: object.generation,
      offset: cursor,
      sha256: await sha256PdfBytes(object.bytes),
    }));
    parts.push(object.bytes, ascii("\n"));
    cursor += object.bytes.length + 1;
  }
  parts.push(incrementalXrefSection(parsed.current, rows, cursor));
  return Object.freeze({
    candidateBytes: concatBytes(parts),
    incrementalXrefOffset: cursor,
    objects: Object.freeze(prepared),
  });
}





















/** The catalog's XMP packet as the source holds it, which the revision
 * writes again without its PDF/A identification. */
export interface PdfNativePdfaIdentificationRemoval {
  readonly objectNumber: number;
  readonly generation: number;
  /** The Metadata object's bytes as the source holds them. */
  readonly expectedObjectSha256: string;
}

export interface PdfNativePreparedPdfaIdentificationRemoval
  extends PdfNativePdfaIdentificationRemoval {
  /** Where the revision wrote it, and the digest of what it wrote. */
  readonly offset: number;
  readonly sha256: string;
}

export const MRAI_NATIVE_PDFA_IDENTIFICATION_MESSAGE =
  "This PDF claims PDF/A conformance in a form the editor cannot remove, so it is not edited.";

const PDFA_IDENTIFICATION_PROPERTIES = "part|amd|conformance|corr|rev";
const PDFA_XML_SPACE = "[ \\t\\r\\n]";
const PDFA_XML_PREFIX = "[A-Za-z_][A-Za-z0-9_.-]*";
const PDFA_IDENTIFICATION_NAMESPACE =
  "(?:\"http://www\\.aiim\\.org/pdfa/ns/id/\"|'http://www\\.aiim\\.org/pdfa/ns/id/')";
const PDFA_IDENTIFICATION_NAMESPACE_TEXT = /pdfaid:|www\.aiim\.org\/pdfa\/ns\/id/i;
const PDFA_WRITABLE_METADATA_ENCODINGS = new Set(["identity", "flate", "ascii85-flate", "ascii-hex", "lzw"]);

function escapePdfaPrefix(prefix: string): string {
  return prefix.replace(/[.\-]/g, (character) => `\\${character}`);
}

/**
 * Where an XMP packet (read one byte per character) states the PDF/A
 * identification: every property of the identification schema, as an
 * element of a prefix bound to it (`<pdfaid:part>1</pdfaid:part>`, or empty),
 * as an attribute of a start tag (` pdfaid:part="2"`, with the space before
 * it), or unprefixed under the schema bound as the default namespace. Null
 * when two such places overlap (a packet this reader does not take). Case is
 * ignored as the save policy's reader ignores it. Mirror of the server's
 * pdfa_identification_ranges.
 */
export function readMraiNativePdfaIdentificationRanges(text: string): Readonly<{
  ranges: readonly (readonly [number, number])[];
  claims: boolean;
}> | null {
  const S = PDFA_XML_SPACE;
  const found: { start: number; end: number; part: boolean }[] = [];
  const elementPatterns: string[] = [];
  const prefixes = new Set<string>();
  for (const match of text.matchAll(
    new RegExp(`xmlns:(${PDFA_XML_PREFIX})${S}*=${S}*${PDFA_IDENTIFICATION_NAMESPACE}`, "gi"),
  )) {
    prefixes.add(match[1]);
  }
  for (const prefix of prefixes) {
    const name = `${escapePdfaPrefix(prefix)}:`;
    elementPatterns.push(name);
    const attribute = new RegExp(
      `${S}+${name}(${PDFA_IDENTIFICATION_PROPERTIES})${S}*=${S}*(?:"[^"]*"|'[^']*')`,
      "gi",
    );
    for (const tag of text.matchAll(/<[A-Za-z_][^<>]*>/g)) {
      for (const match of tag[0].matchAll(attribute)) {
        const start = tag.index! + match.index!;
        found.push({ start, end: start + match[0].length, part: match[1].toLowerCase() === "part" });
      }
    }
  }
  if (new RegExp(`xmlns${S}*=${S}*${PDFA_IDENTIFICATION_NAMESPACE}`, "i").test(text)) {
    elementPatterns.push("");
  }
  for (const name of elementPatterns) {
    for (const pattern of [
      `<${name}(${PDFA_IDENTIFICATION_PROPERTIES})(?:${S}[^<>]*)?>[^<]*</${name}\\1${S}*>`,
      `<${name}(${PDFA_IDENTIFICATION_PROPERTIES})(?:${S}[^<>]*)?/>`,
    ]) {
      for (const match of text.matchAll(new RegExp(pattern, "gi"))) {
        found.push({
          start: match.index!,
          end: match.index! + match[0].length,
          part: match[1].toLowerCase() === "part",
        });
      }
    }
  }
  found.sort((left, right) => left.start - right.start || left.end - right.end);
  const ranges: (readonly [number, number])[] = [];
  for (const range of found) {
    const last = ranges.at(-1);
    if (last && last[0] === range.start && last[1] === range.end) continue;
    if (last && range.start < last[1]) return null;
    ranges.push(Object.freeze([range.start, range.end] as const));
  }
  return Object.freeze({ ranges: Object.freeze(ranges), claims: found.some((range) => range.part) });
}

/** The packet with each range replaced by as many spaces. */
function blankedPdfaIdentification(
  bytes: Uint8Array,
  ranges: readonly (readonly [number, number])[],
): Uint8Array {
  const blanked = Uint8Array.from(bytes);
  for (const [start, end] of ranges) blanked.fill(0x20, start, end);
  return blanked;
}

/** The Metadata object the catalog names, when it names one by reference. */
function catalogMetadataReference(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
): Readonly<{ objectNumber: number; generation: number }> | null {
  const catalog = parseIndirectObject(
    sourceBytes, parsed, parsed.current.rootObjectNumber, parsed.current.rootGeneration, undefined, "opaque",
  );
  const catalogSource = asLatin1(catalog.bytes);
  const entry = oneTopLevelDictionaryEntry(catalogSource, "Metadata", "PDF Catalog Metadata", true);
  if (!entry) return null;
  const value = /^([0-9]{1,10})[\x00\t\n\f\r ]+([0-9]{1,5})[\x00\t\n\f\r ]+R$/.exec(
    catalogSource.slice(entry.valueStart, entry.valueEnd).replace(/^[\x00\t\n\f\r ]+|[\x00\t\n\f\r ]+$/g, ""),
  );
  if (!value) return null;
  const reference = { objectNumber: Number(value[1]), generation: Number(value[2]) };
  assertObjectReference(reference.objectNumber, reference.generation);
  return reference;
}

/**
 * The catalog's Metadata object written again without its PDF/A
 * identification, or null when the source claims no PDF/A conformance.
 * Refused (`unsupported-document`) when the source claims it, or may, in a
 * form this cannot write: an undecodable or object-stream packet that names
 * the identification schema, overlapping properties, a catalog this reader
 * cannot read in a file that names the schema. Mirror of the server's
 * pdfa_identification_revision.
 */
async function readPdfaIdentificationRevision(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
): Promise<Readonly<{
  removal: PdfNativePdfaIdentificationRemoval;
  bytes: Uint8Array;
}> | null> {
  const refuse = (): never => fail("unsupported-document", MRAI_NATIVE_PDFA_IDENTIFICATION_MESSAGE);
  let reference: Readonly<{ objectNumber: number; generation: number }> | null;
  let object: ParsedIndirectObject | null = null;
  try {
    reference = catalogMetadataReference(sourceBytes, parsed);
    if (reference) {
      object = parseIndirectObject(
        sourceBytes, parsed, reference.objectNumber, reference.generation, undefined, "opaque",
      );
    }
  } catch (error) {
    if (!(error instanceof PdfNativeTransactionError)) throw error;
    // Neither the catalog nor its packet can be read: a file that names the
    // identification schema anywhere may claim PDF/A and is not written.
    if (PDFA_IDENTIFICATION_NAMESPACE_TEXT.test(decodedSourceLatin1(sourceBytes))) refuse();
    return null;
  }
  if (!reference || !object) return null;
  if (
    object.storage !== "direct" ||
    object.streamStart === null ||
    object.streamBytes === null ||
    !PDFA_WRITABLE_METADATA_ENCODINGS.has(object.streamEncoding ?? "")
  ) {
    const readable = object.streamBytes ?? object.encodedStreamBytes ?? object.bytes;
    if (PDFA_IDENTIFICATION_NAMESPACE_TEXT.test(asLatin1(readable))) refuse();
    return null;
  }
  const identification = readMraiNativePdfaIdentificationRanges(asLatin1(object.streamBytes));
  if (!identification) refuse();
  if (!identification!.claims) return null;
  const blanked = blankedPdfaIdentification(object.streamBytes, identification!.ranges);
  assertDirectlyWritableStream(object);
  const encoded = encodeLikeTarget(
    object.streamEncoding as "identity" | "flate" | "ascii85-flate" | "ascii-hex" | "lzw",
    blanked,
  );
  return Object.freeze({
    removal: Object.freeze({
      objectNumber: reference.objectNumber,
      generation: reference.generation,
      expectedObjectSha256: await sha256PdfBytes(object.bytes),
    }),
    bytes: rewriteStreamObjectBytes(sourceBytes, object, encoded),
  });
}

/** Whether the source claims PDF/A conformance (the revision a write over
 * it appends drops the claim). */
export async function readMraiNativePdfaIdentificationRemoval(
  sourceBytes: Uint8Array,
): Promise<PdfNativePdfaIdentificationRemoval | null> {
  return (await readPdfaIdentificationRevision(sourceBytes, parsePdf(sourceBytes, false)))?.removal ?? null;
}

/**
 * `base` (one revision appended to the source) with the catalog's Metadata
 * written last in that revision without its PDF/A identification, or null
 * when the source claims no PDF/A conformance. `expected` is what a record
 * planned earlier names: the source must still hold that packet.
 */
async function appendPdfaIdentificationRemoval(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
  base: Readonly<{ candidateBytes: Uint8Array; incrementalXrefOffset: number }>,
): Promise<Readonly<{
  candidateBytes: Uint8Array;
  incrementalXrefOffset: number;
  removal: PdfNativePreparedPdfaIdentificationRemoval;
}> | null> {
  const revision = await readPdfaIdentificationRevision(sourceBytes, parsed);
  if (!revision) return null;
  const rows = appendedRevisionRows(parsed, base);
  if (rows.some((row) => row.objectNumber === revision.removal.objectNumber)) {
    fail("invalid-request", "The catalog's Metadata is an object the write revises.");
  }
  const appended = await appendObjectsToRevision(parsed, base, rows, [{
    objectNumber: revision.removal.objectNumber,
    generation: revision.removal.generation,
    bytes: revision.bytes,
  }]);
  return Object.freeze({
    candidateBytes: appended.candidateBytes,
    incrementalXrefOffset: appended.incrementalXrefOffset,
    removal: Object.freeze({
      ...revision.removal,
      offset: appended.objects[0].offset,
      sha256: appended.objects[0].sha256,
    }),
  });
}

/** A revision another writer built (a picture replacement), with the PDF/A
 * identification dropped in it when the source claims PDF/A. */
export async function appendMraiNativePdfaIdentificationRemoval(
  sourceBytes: Uint8Array,
  base: Readonly<{ candidateBytes: Uint8Array; incrementalXrefOffset: number }>,
): Promise<Readonly<{
  candidateBytes: Uint8Array;
  incrementalXrefOffset: number;
  removal: PdfNativePreparedPdfaIdentificationRemoval;
}> | null> {
  return appendPdfaIdentificationRemoval(sourceBytes, parsePdf(sourceBytes, false), base);
}










export interface PdfNativeCatalogVersionRaise {
  readonly objectNumber: number;
  readonly generation: number;
  /** The catalog's bytes as the source holds them. */
  readonly expectedObjectSha256: string;
  readonly version: "1.4";
}

export interface PdfNativePreparedCatalogVersionRaise extends PdfNativeCatalogVersionRaise {
  readonly offset: number;
  readonly sha256: string;
}

/** Whether what a revision writes (stream data aside) draws with
 * transparency, PDF 1.4: constant alpha, a blend mode, a soft mask, a
 * transparency group. A widget's /MK /CA is a caption string, not alpha. */
function revisionDrawsTransparency(written: Uint8Array): boolean {
  const text = asLatin1(written).replace(/stream(?:\r\n|\n|\r)[\s\S]*?endstream/g, "stream endstream");
  return /\/(?:SMask|BM|AIS|TK)(?![^\x00\t\n\f\r ()<>[\]{}/%])/.test(text) ||
    /\/(?:ca|CA)[\x00\t\n\f\r ]*[-+]?(?:\d|\.\d)/.test(text) ||
    /\/S[\x00\t\n\f\r ]*\/Transparency(?![^\x00\t\n\f\r ()<>[\]{}/%])/.test(text);
}

/**
 * `base` (one revision appended to the source, its PDF/A packet included)
 * with the catalog written last stating PDF 1.4, or null when the revision
 * draws nothing 1.4 needs, the file already declares 1.4 or later, or its
 * catalog is not one this can write (it names a version already, sits in an
 * object stream, or is an object the revision writes). Never refuses: a
 * file this cannot state a version in is written as before.
 */
async function appendCatalogVersionRaise(
  sourceBytes: Uint8Array,
  parsed: ParsedPdf,
  base: Readonly<{ candidateBytes: Uint8Array; incrementalXrefOffset: number }>,
): Promise<Readonly<{
  candidateBytes: Uint8Array;
  incrementalXrefOffset: number;
  raise: PdfNativePreparedCatalogVersionRaise;
}> | null> {
  const header = /%PDF-(\d)\.(\d)/.exec(asLatin1(sourceBytes.subarray(0, 1024)));
  if (!header) return null;
  const declared = [Number(header[1]), Number(header[2])] as const;
  if (declared[0] > 1 || (declared[0] === 1 && declared[1] >= 4)) return null;
  if (!revisionDrawsTransparency(base.candidateBytes.subarray(sourceBytes.length, base.incrementalXrefOffset))) {
    return null;
  }
  let catalog: ParsedIndirectObject;
  let rows: { objectNumber: number; generation: number; offset: number }[];
  try {
    catalog = parseIndirectObject(
      sourceBytes, parsed, parsed.current.rootObjectNumber, parsed.current.rootGeneration, undefined, "opaque",
    );
    if (catalog.storage !== "direct" || catalog.dictionaryStart === null) return null;
    if (oneTopLevelDictionaryEntry(asLatin1(catalog.bytes), "Version", "PDF Catalog Version", true)) return null;
    rows = appendedRevisionRows(parsed, base);
  } catch (error) {
    if (error instanceof PdfNativeTransactionError) return null;
    throw error;
  }
  if (rows.some((row) => row.objectNumber === catalog.objectNumber)) return null;
  const inside = catalog.dictionaryStart - catalog.start + 2;
  if (asLatin1(catalog.bytes.subarray(inside - 2, inside)) !== "<<") return null;
  const written = concatBytes([
    catalog.bytes.subarray(0, inside),
    ascii(" /Version /1.4 "),
    catalog.bytes.subarray(inside),
  ]);
  const appended = await appendObjectsToRevision(parsed, base, rows, [{
    objectNumber: catalog.objectNumber,
    generation: catalog.generation,
    bytes: written,
  }]);
  return Object.freeze({
    candidateBytes: appended.candidateBytes,
    incrementalXrefOffset: appended.incrementalXrefOffset,
    raise: Object.freeze({
      objectNumber: catalog.objectNumber,
      generation: catalog.generation,
      expectedObjectSha256: await sha256PdfBytes(catalog.bytes),
      version: "1.4" as const,
      offset: appended.objects[0].offset,
      sha256: appended.objects[0].sha256,
    }),
  });
}


/**
 * A history revision that swaps whole objects (a font extension's Undo and
 * Redo) with the structure elements its origin wrote written back after
 * them (appendStructActualTextRevisions).
 */
export async function appendMraiNativeStructActualTextRevisions(
  sourceBytes: Uint8Array,
  base: Readonly<{ candidateBytes: Uint8Array; incrementalXrefOffset: number }>,
  revisions: readonly PdfNativeStructActualTextRevision[],
): Promise<Readonly<{
  candidateBytes: Uint8Array;
  incrementalXrefOffset: number;
  objects: readonly Readonly<{ objectNumber: number; generation: number; offset: number; sha256: string }>[];
}>> {
  return appendStructActualTextRevisions(sourceBytes, parsePdf(sourceBytes, false), base, revisions);
}

/**
 * The decoded stream a request writes: its replacement over the operator
 * span, the same bytes (or their part) over each duplicate copy, and each
 * text rule's new bytes over the rule. The request's preconditions are the
 * caller's to check (prepare checks them before it writes).
 */
export function applyMraiNativeRequestSpans(
  streamBytes: Uint8Array,
  request: Pick<
    PdfNativeContentStreamTransactionRequest,
    "operatorSpan" | "replacement" | "duplicateSpans" | "textRules" | "boxMark"
  >,
): Uint8Array {
  // A box mark's halves are written round the box: the opening before the
  // span it starts at, the closing after the one it ends at (in this order
  // the stable sort keeps them, mirror of replay).
  const mark = request.boxMark;
  const editedSpans = [
    ...(mark ? [{ start: mark.start, end: mark.start, bytes: MRAI_NATIVE_BOX_MARK_OPEN }] : []),
    { start: request.operatorSpan.start, end: request.operatorSpan.end, bytes: request.replacement },
    ...(request.duplicateSpans ?? []).map((span) => ({
      start: span.start,
      end: span.end,
      bytes: span.replacementStart !== undefined
        ? request.replacement.slice(span.replacementStart, span.replacementEnd)
        : request.replacement,
    })),
    ...(request.textRules ?? []).map((rule) => ({ start: rule.start, end: rule.end, bytes: rule.added })),
    ...(mark ? [{ start: mark.end, end: mark.end, bytes: MRAI_NATIVE_BOX_MARK_CLOSE }] : []),
  ].sort((left, right) => left.start - right.start);
  const changedParts: Uint8Array[] = [];
  let cursor = 0;
  for (const span of editedSpans) {
    if (span.start < cursor || span.end > streamBytes.length || span.end < span.start) {
      fail("invalid-request", "The request's spans overlap or leave the stream.");
    }
    changedParts.push(streamBytes.slice(cursor, span.start), span.bytes);
    cursor = span.end;
  }
  changedParts.push(streamBytes.slice(cursor));
  return concatBytes(changedParts);
}

/**
 * A content-stream transaction prepared once for the same request on the same
 * bytes: the editor's exact preview of an edit and its write when the line is
 * left asked the same Worker the same question (the run's format plan they
 * already share: list-text-fidelity propertyPlans), each preparing and then
 * reopening the whole candidate again (0.035 s and 0.05 s on the 5 MB IPCC
 * report). Asked again within a minute, the answer is that preparation, a copy
 * to each asker (a caller zeroes what it is handed); a refusal is not kept.
 * The question is the request with every byte array named by its digest, and
 * the source by its digest; a request holding anything JSON does not spell
 * out exactly (a Map, a Set, a class) is never shared.
 */
const PREPARED_TRANSACTIONS_KEPT = 4;
const PREPARED_TRANSACTION_KEPT_MS = 60_000;
const preparedTransactions = new Map<string, { at: number; prepared: Promise<PdfNativePreparedContentStreamTransaction> }>();
/** The verifications of prepared candidates, by everything they read. */
const verifiedCandidates = new Map<string, { at: number; verification: PdfNativeReopenVerification }>();
let preparedAnswersShared = 0;
/** How many asks were answered with a preparation or a verification already
 * made (a probe reads it around a write). */
export function pdfNativePreparedAnswersShared(): number {
  return preparedAnswersShared;
}

/** JSON with every object's keys in one order and every byte array named by
 * its length and digest; null when the value holds anything else JSON would
 * not spell out exactly. `skip` (a byte array) is left out. */
async function byteNamedFingerprint(value: unknown, skip?: Uint8Array): Promise<string | null> {
  const views = new Map<ArrayBufferView, string>();
  let plain = true;
  const seen = new Set<unknown>();
  const walk = (held: unknown) => {
    // (A function or a symbol JSON leaves out: never one question.)
    if (typeof held === "function" || typeof held === "symbol" || typeof held === "bigint") { plain = false; return; }
    if (!plain || held === null || typeof held !== "object" || held === skip || seen.has(held)) return;
    seen.add(held);
    if (ArrayBuffer.isView(held)) { views.set(held, ""); return; }
    const prototype = Object.getPrototypeOf(held);
    if (!Array.isArray(held) && prototype !== Object.prototype && prototype !== null) { plain = false; return; }
    for (const child of Array.isArray(held) ? held : Object.values(held as Record<string, unknown>)) walk(child);
  };
  walk(value);
  if (!plain) return null;
  for (const view of views.keys()) {
    views.set(view, await sha256PdfBytes(new Uint8Array(view.buffer, view.byteOffset, view.byteLength)));
  }
  return JSON.stringify(value, (_key, held: unknown) => {
    if (held === skip) return "<skipped>";
    if (held !== null && typeof held === "object" && ArrayBuffer.isView(held)) {
      return `bytes:${held.byteLength}:${views.get(held) ?? "?"}`;
    }
    if (held && typeof held === "object" && !Array.isArray(held)) {
      return Object.fromEntries(Object.entries(held as Record<string, unknown>)
        .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0)));
    }
    if (typeof held === "number" && !Number.isFinite(held)) return `number:${String(held)}`;
    return held;
  });
}

function prunePreparedCaches(now: number): void {
  for (const [key, entry] of preparedTransactions) if (now - entry.at > PREPARED_TRANSACTION_KEPT_MS) preparedTransactions.delete(key);
  for (const [key, entry] of verifiedCandidates) if (now - entry.at > PREPARED_TRANSACTION_KEPT_MS) verifiedCandidates.delete(key);
  while (preparedTransactions.size > PREPARED_TRANSACTIONS_KEPT) preparedTransactions.delete(preparedTransactions.keys().next().value!);
  while (verifiedCandidates.size > PREPARED_TRANSACTIONS_KEPT) verifiedCandidates.delete(verifiedCandidates.keys().next().value!);
}

export async function prepareMraiNativeContentStreamTransaction(
  sourceBytes: Uint8Array,
  rawRequest: PdfNativeContentStreamTransactionRequest,
): Promise<PdfNativePreparedContentStreamTransaction> {
  const question = await byteNamedFingerprint(rawRequest);
  if (question === null) return prepareMraiNativeContentStreamTransactionOnce(sourceBytes, rawRequest);
  const key = JSON.stringify([await sha256PdfBytes(sourceBytes), sourceBytes.length, question]);
  const now = Date.now();
  prunePreparedCaches(now);
  const entry = preparedTransactions.get(key);
  if (entry) {
    const kept = await entry.prepared.catch(() => null);
    const copy = kept && copyPlainPrepared(kept);
    if (copy) {
      preparedAnswersShared += 1;
      return copy as PdfNativePreparedContentStreamTransaction;
    }
  }
  // Prepared for this asker on its own bytes, exactly as it always was; a
  // copy is kept (taken before the asker can zero what it is handed).
  const prepared = prepareMraiNativeContentStreamTransactionOnce(sourceBytes, rawRequest);
  const settled = { at: now, prepared: prepared.then((answer) => {
    const kept = copyPlainPrepared(answer);
    if (!kept) throw new Error("Not kept.");
    return kept as PdfNativePreparedContentStreamTransaction;
  }) };
  preparedTransactions.set(key, settled);
  // A refusal (or an answer that cannot be copied) is not kept.
  settled.prepared.catch(() => { if (preparedTransactions.get(key) === settled) preparedTransactions.delete(key); });
  prunePreparedCaches(now);
  return prepared;
}

/** A deep copy of plain objects, arrays and byte arrays (each byte array of
 * its own constructor), frozen where the original was; null when the value
 * holds anything else. */
function copyPlainPrepared(value: unknown, seen = new Map<unknown, unknown>()): unknown {
  if (value === null || typeof value !== "object") return value;
  if (seen.has(value)) return seen.get(value);
  if (ArrayBuffer.isView(value)) {
    const copy = (value as unknown as { slice(): unknown }).slice();
    seen.set(value, copy);
    return copy;
  }
  const prototype = Object.getPrototypeOf(value);
  const isArray = Array.isArray(value);
  if (!isArray && prototype !== Object.prototype && prototype !== null) return null;
  const copy: Record<string, unknown> | unknown[] = isArray ? [] : {};
  seen.set(value, copy);
  for (const [name, held] of Object.entries(value as Record<string, unknown>)) {
    const copied = copyPlainPrepared(held, seen);
    if (copied === null && held !== null) return null;
    (copy as Record<string, unknown>)[name] = copied;
  }
  return Object.isFrozen(value) ? Object.freeze(copy) : copy;
}

async function prepareMraiNativeContentStreamTransactionOnce(
  sourceBytes: Uint8Array,
  rawRequest: PdfNativeContentStreamTransactionRequest,
): Promise<PdfNativePreparedContentStreamTransaction> {
  const request = normalizeRequest(rawRequest);
  const sourceSha256 = await sha256PdfBytes(sourceBytes);
  if (sourceSha256 !== request.identity.sourceSha256) {
    fail("source-mismatch", "Loaded PDF bytes do not match the pinned source.");
  }
  const parsed = parsePdf(sourceBytes, false);
  // A consolidated page edits the virtual joined stream; every other
  // transaction edits an object the file already holds.
  const consolidated = request.pageContentConsolidation
    ? await consolidatedVirtualTarget(sourceBytes, parsed, request.pageContentConsolidation)
    : null;
  if (
    consolidated &&
    (request.identity.objectNumber !== request.pageContentConsolidation!.stream.objectNumber ||
      request.identity.generation !== 0)
  ) {
    fail("invalid-request", "Consolidated content stream identity is inconsistent.");
  }
  const object = consolidated
    ? consolidated.object
    : parseIndirectObject(
        sourceBytes,
        parsed,
        request.identity.objectNumber,
        request.identity.generation,
      );
  assertTargetStream(object);

  if ((await sha256PdfBytes(object.bytes)) !== request.expectedObjectSha256) {
    fail("precondition-mismatch", "Target PDF object precondition changed.");
  }
  if (
    (await sha256PdfBytes(object.encodedStreamBytes)) !==
    request.expectedStreamSha256
  ) {
    fail(
      "precondition-mismatch",
      "Target content stream precondition changed.",
    );
  }
  if (request.operatorSpan.end > object.streamBytes.length) {
    fail("invalid-request", "Operator span exceeds the target content stream.");
  }
  const originalSpan = object.streamBytes.slice(
    request.operatorSpan.start,
    request.operatorSpan.end,
  );
  if (
    (await sha256PdfBytes(originalSpan)) !== request.operatorSpan.expectedSha256
  ) {
    fail("precondition-mismatch", "Target operator span precondition changed.");
  }
  if (
    request.replacementPieces &&
    !bytesEqual(
      expandPdfNativeReplacementPieces(originalSpan, request.replacementPieces),
      request.replacement,
    )
  ) {
    fail("invalid-request", "Replacement pieces do not put together the replacement.");
  }
  await assertUntouchedPreconditions(
    sourceBytes,
    parsed,
    request.untouchedObjects ?? [],
  );

  // Every duplicate span holds the operator span's bytes (a partial one,
  // the bytes of its source range) and takes the replacement (its part of
  // it) too (mirror of replay).
  const operatorStart = request.operatorSpan.start;
  for (const span of request.duplicateSpans ?? []) {
    const held = span.sourceStart !== undefined
      ? originalSpan.slice(span.sourceStart - operatorStart, span.sourceEnd! - operatorStart)
      : originalSpan;
    if (span.end > object.streamBytes.length ||
      !bytesEqual(object.streamBytes.slice(span.start, span.end), held)) {
      fail("precondition-mismatch", "A duplicate span no longer holds the operator span's bytes.");
    }
  }
  // Each rule the edit rewrites stands where it says, beside the run's text
  // object: an underline ends at its `BT`, a strikethrough starts after its
  // `ET` (mirror of replay).
  if (request.textRules?.length) {
    const edges = pdfNativeTextObjectEdgesAt(LATIN1.decode(object.streamBytes), request.operatorSpan.start);
    for (const rule of request.textRules) {
      const edge = rule.style === "underline" ? rule.end === edges.start : rule.start === edges.end;
      if (!edge || rule.end > object.streamBytes.length ||
        !bytesEqual(object.streamBytes.slice(rule.start, rule.end), rule.removed)) {
        fail("precondition-mismatch", "A text rule is no longer where it was drawn.");
      }
    }
  }
  const changedDecodedStream = applyMraiNativeRequestSpans(object.streamBytes, request);
  if (request.boxMark) {
    const issue = mraiNativeBoxMarkIssue(
      LATIN1.decode(object.streamBytes),
      LATIN1.decode(changedDecodedStream),
      request.boxMark,
      request.operatorSpan,
      request.replacement.length - (request.operatorSpan.end - request.operatorSpan.start),
    );
    if (issue) fail("invalid-request", issue);
  }
  if (!actualTextClustersKeepTheirShows(object.streamBytes, changedDecodedStream,
    request.contentAppend?.direction === "append" ? object.streamBytes.length : undefined)) {
    fail("unsupported-document", MRAI_NATIVE_ACTUAL_TEXT_CLUSTER_MESSAGE);
  }
  const changedStream =
    object.streamEncoding === "flate"
      ? encodeDeterministicFlateStream(changedDecodedStream)
      : object.streamEncoding === "ascii85-flate"
        ? encodeAscii85Stream(
            encodeDeterministicFlateStream(changedDecodedStream),
          )
        : object.streamEncoding === "ascii-hex"
          ? encodeAsciiHexStream(changedDecodedStream)
          : object.streamEncoding === "lzw"
            ? encodeLzwStream(changedDecodedStream)
            : changedDecodedStream;
  const isolated = request.formInstanceIsolation
    ? await prepareFormInstanceIsolation(
        sourceBytes,
        parsed,
        request,
        object,
        changedDecodedStream,
        changedStream,
      )
    : null;
  const pageExtGState = request.pageExtGStateMutation
    ? await preparePageExtGStateMutation(
        sourceBytes,
        parsed,
        request,
        object,
        changedDecodedStream,
        changedStream,
      )
    : null;
  const pageXObjectUnmap = request.pageXObjectUnmap
    ? await preparePageXObjectUnmap(
        sourceBytes,
        parsed,
        request,
        object,
        changedDecodedStream,
        changedStream,
      )
    : null;
  const pageFontExtension =
    request.pageFontExtensionMutation && !request.formInstanceIsolation
      ? request.pageFontExtensionMutation.form
        ? await prepareFormFontExtensionMutation(
            sourceBytes,
            parsed,
            request,
            object,
            changedDecodedStream,
            changedStream,
          )
        : await preparePageFontExtensionMutation(
            sourceBytes,
            parsed,
            request,
            object,
            changedDecodedStream,
            changedStream,
            consolidated?.bytes ?? null,
          )
      : null;
  const contentAppend = request.contentAppend
    ? await prepareContentAppend(
        sourceBytes,
        parsed,
        request,
        object,
        changedDecodedStream,
        changedStream,
      )
    : null;
  const pageFontReference = request.pageFontReference
    ? await preparePageFontReference(
        sourceBytes,
        parsed,
        request,
        object,
        changedDecodedStream,
        changedStream,
      )
    : null;
  if (consolidated && (isolated || pageExtGState || pageXObjectUnmap || contentAppend || pageFontReference)) {
    fail("invalid-request", "Page content consolidation cannot share another mutation transaction.");
  }
  // A font extension on a consolidated page writes the stream and the page
  // itself; only a plain edit takes the consolidation route alone.
  const consolidation = consolidated && !pageFontExtension
    ? await preparePageContentConsolidation(
        sourceBytes,
        parsed,
        request,
        object,
        consolidated.bytes,
        changedDecodedStream,
        changedStream,
      )
    : null;
  const specialized =
    isolated ?? pageExtGState ?? pageXObjectUnmap ?? pageFontExtension ??
    contentAppend ?? pageFontReference ?? consolidation;
  // A consolidation rides on its own preparation or on a page font
  // extension; a form font extension refuses one and carries none.
  const pageContentConsolidation =
    consolidation?.pageContentConsolidation ?? pageFontExtension?.pageContentConsolidation;
  const revisedObject = specialized
    ? null
    : revisedContentStreamObject(sourceBytes, object, changedStream);
  const incremental = specialized
    ? null
    : makeIncrementalRevision(
        sourceBytes,
        parsed.current,
        request.identity.objectNumber,
        request.identity.generation,
        revisedObject!,
      );
  // The structure elements whose /ActualText names this text are written in
  // the same revision, after the content write's own objects.
  const structActualText = request.structActualText
    ? await appendStructActualTextRevisions(
        sourceBytes,
        parsed,
        {
          candidateBytes: specialized?.candidateBytes ?? incremental!.candidateBytes,
          incrementalXrefOffset: specialized?.incrementalXrefOffset ?? incremental!.incrementalXrefOffset,
        },
        request.structActualText,
      )
    : null;
  // A source that claims PDF/A conformance no longer does once written: its
  // catalog's XMP packet is written last in the same revision without the
  
  const pdfaIdentification = await appendPdfaIdentificationRemoval(sourceBytes, parsed, {
    candidateBytes: structActualText?.candidateBytes ?? specialized?.candidateBytes ?? incremental!.candidateBytes,
    incrementalXrefOffset:
      structActualText?.incrementalXrefOffset ?? specialized?.incrementalXrefOffset ?? incremental!.incrementalXrefOffset,
  });
  // A shape's opacity or blend, or a picture with an alpha channel, in a
  // file declaring PDF 1.3: the catalog states 1.4, last in the same
  // revision (user audit PDF 0020).
  const catalogVersion = await appendCatalogVersionRaise(sourceBytes, parsed, {
    candidateBytes:
      pdfaIdentification?.candidateBytes ??
      structActualText?.candidateBytes ?? specialized?.candidateBytes ?? incremental!.candidateBytes,
    incrementalXrefOffset:
      pdfaIdentification?.incrementalXrefOffset ??
      structActualText?.incrementalXrefOffset ??
      specialized?.incrementalXrefOffset ?? incremental!.incrementalXrefOffset,
  });
  const candidateBytes =
    catalogVersion?.candidateBytes ??
    pdfaIdentification?.candidateBytes ??
    structActualText?.candidateBytes ?? specialized?.candidateBytes ?? incremental!.candidateBytes;
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_CONTENT_STREAM_SCHEMA_VERSION,
    request,
    sourceLength: sourceBytes.length,
    sourceXrefOffset: parsed.current.offset,
    sourceXrefSections: parsed.sections.length,
    changedObjectOffset:
      specialized?.changedObjectOffset ?? incremental!.changedObjectOffset,
    incrementalXrefOffset:
      catalogVersion?.incrementalXrefOffset ??
      pdfaIdentification?.incrementalXrefOffset ??
      structActualText?.incrementalXrefOffset ??
      specialized?.incrementalXrefOffset ?? incremental!.incrementalXrefOffset,
    changedObjectSha256:
      specialized?.changedObjectSha256 ??
      (await sha256PdfBytes(revisedObject!)),
    changedStreamSha256:
      specialized?.changedStreamSha256 ?? (await sha256PdfBytes(changedStream)),
    changedDecodedStreamSha256:
      specialized?.changedDecodedStreamSha256 ??
      (await sha256PdfBytes(changedDecodedStream)),
    replacementSha256: await sha256PdfBytes(request.replacement),
    candidateSha256: await sha256PdfBytes(candidateBytes),
    candidateBytes,
    ...(isolated
      ? { formInstanceIsolation: isolated.formInstanceIsolation }
      : {}),
    ...(pageExtGState
      ? { pageExtGStateMutation: pageExtGState.pageExtGStateMutation }
      : {}),
    ...(pageXObjectUnmap
      ? { pageXObjectUnmap: pageXObjectUnmap.pageXObjectUnmap }
      : {}),
    ...((pageFontExtension ?? isolated)?.pageFontExtensionMutation
      ? {
          pageFontExtensionMutation: (pageFontExtension ?? isolated)!
            .pageFontExtensionMutation!,
        }
      : {}),
    ...(contentAppend ? { contentAppend: contentAppend.contentAppend } : {}),
    ...(pageFontReference ? { pageFontReference: pageFontReference.pageFontReference } : {}),
    ...(pageContentConsolidation ? { pageContentConsolidation } : {}),
    ...(structActualText ? { structActualText: structActualText.objects } : {}),
    ...(pdfaIdentification ? { pdfaIdentificationRemoval: pdfaIdentification.removal } : {}),
    ...(catalogVersion ? { catalogVersionRaise: catalogVersion.raise } : {}),
  });
}

export async function verifyMraiNativeContentStreamCandidate(
  sourceBytes: Uint8Array,
  prepared: PdfNativePreparedContentStreamTransaction,
): Promise<PdfNativeReopenVerification> {
  // The reopening reads only the source, the candidate and what the
  // preparation says of it: all three as they were when one passed, this one
  // passes (the preview's and the write's of one edit, see prepare above). The
  // candidate's prefix is compared with the source every time.
  const said = await byteNamedFingerprint(prepared, prepared.candidateBytes);
  const key = said === null ? null : JSON.stringify([await sha256PdfBytes(sourceBytes), sourceBytes.length,
    await sha256PdfBytes(prepared.candidateBytes), prepared.candidateBytes.length, said]);
  const now = Date.now();
  prunePreparedCaches(now);
  const known = key === null ? undefined : verifiedCandidates.get(key);
  if (known) {
    if (prepared.sourceLength !== sourceBytes.length ||
      !bytesEqual(prepared.candidateBytes.subarray(0, sourceBytes.length), sourceBytes)) {
      fail("verification-failed", "Original PDF bytes are not an exact prefix.");
    }
    candidateSourcePrefixes.set(prepared.candidateBytes, sourceBytes);
    preparedAnswersShared += 1;
    return Object.freeze({ ...known.verification, identity: prepared.request.identity });
  }
  const verification = await verifyMraiNativeContentStreamCandidateOnce(sourceBytes, prepared);
  if (key !== null) {
    verifiedCandidates.set(key, { at: now, verification });
    prunePreparedCaches(now);
  }
  return verification;
}

async function verifyMraiNativeContentStreamCandidateOnce(
  sourceBytes: Uint8Array,
  prepared: PdfNativePreparedContentStreamTransaction,
): Promise<PdfNativeReopenVerification> {
  if (prepared.sourceLength !== sourceBytes.length) {
    fail("verification-failed", "Candidate source length changed.");
  }
  const prefix = prepared.candidateBytes.subarray(0, sourceBytes.length);
  if (!bytesEqual(prefix, sourceBytes)) {
    fail("verification-failed", "Original PDF bytes are not an exact prefix.");
  }
  candidateSourcePrefixes.set(prepared.candidateBytes, sourceBytes);
  if (
    (await sha256PdfBytes(sourceBytes)) !==
    prepared.request.identity.sourceSha256
  ) {
    fail("verification-failed", "Reopen source identity changed.");
  }
  if (
    (await sha256PdfBytes(prepared.candidateBytes)) !== prepared.candidateSha256
  ) {
    fail("verification-failed", "Candidate PDF hash changed.");
  }

  const reopened = parsePdf(prepared.candidateBytes, true);
  if (
    reopened.current.offset !== prepared.incrementalXrefOffset ||
    reopened.current.previousOffset !== prepared.sourceXrefOffset ||
    reopened.previous?.offset !== prepared.sourceXrefOffset
  ) {
    fail("verification-failed", "Incremental xref/trailer linkage is invalid.");
  }
  const isolation = prepared.formInstanceIsolation;
  if (isolation) {
    for (const item of isolation.objects) {
      const entry = reopened.current.entries.get(item.objectNumber);
      if (
        !entry?.inUse ||
        entry.generation !== item.generation ||
        entry.offset !== item.offset
      ) {
        fail(
          "verification-failed",
          "Isolated Form incremental ownership is invalid.",
        );
      }
      const reopenedObject = parseIndirectObject(
        prepared.candidateBytes,
        reopened,
        item.objectNumber,
        item.generation,
      );
      if ((await sha256PdfBytes(reopenedObject.bytes)) !== item.sha256) {
        fail(
          "verification-failed",
          "Isolated Form object failed deterministic reopen verification.",
        );
      }
    }
    const original = parseIndirectObject(
      prepared.candidateBytes,
      reopened,
      prepared.request.identity.objectNumber,
      prepared.request.identity.generation,
    );
    if (
      (await sha256PdfBytes(original.bytes)) !==
        prepared.request.expectedObjectSha256 ||
      !original.encodedStreamBytes ||
      (await sha256PdfBytes(original.encodedStreamBytes)) !==
        prepared.request.expectedStreamSha256
    ) {
      fail(
        "verification-failed",
        "Original shared Form changed during instance isolation.",
      );
    }
  } else {
    const currentEntry = reopened.current.entries.get(
      prepared.request.identity.objectNumber,
    );
    if (
      !currentEntry?.inUse ||
      currentEntry.generation !== prepared.request.identity.generation ||
      currentEntry.offset !== prepared.changedObjectOffset
    ) {
      fail("verification-failed", "Incremental object ownership is invalid.");
    }
  }
  const pageExtGState = prepared.pageExtGStateMutation;
  const pageExtGStateRequest = prepared.request.pageExtGStateMutation;
  if (Boolean(pageExtGState) !== Boolean(pageExtGStateRequest)) {
    fail("verification-failed", "Page ExtGState receipt is inconsistent.");
  }
  if (pageExtGState && pageExtGStateRequest) {
    if (
      pageExtGState.resourceName !== pageExtGStateRequest.resourceName ||
      pageExtGState.reused !== pageExtGStateRequest.reuseExisting
    ) {
      fail("verification-failed", "Page ExtGState receipt identity changed.");
    }
    if (!pageExtGState.reused) {
      for (const item of [
        pageExtGState.graphicsState,
        pageExtGState.resources,
        pageExtGState.page,
      ]) {
        const entry = reopened.current.entries.get(item.objectNumber);
        if (
          !entry?.inUse ||
          entry.generation !== item.generation ||
          entry.offset !== item.offset
        ) {
          fail(
            "verification-failed",
            "Page ExtGState incremental ownership is invalid.",
          );
        }
        const object = parseIndirectObject(
          prepared.candidateBytes,
          reopened,
          item.objectNumber,
          item.generation,
        );
        if ((await sha256PdfBytes(object.bytes)) !== item.sha256) {
          fail(
            "verification-failed",
            "Page ExtGState object failed deterministic reopen verification.",
          );
        }
      }
      if (
        pageExtGState.page.objectNumber !==
          pageExtGStateRequest.page.objectNumber ||
        pageExtGState.page.generation !== pageExtGStateRequest.page.generation
      ) {
        fail("verification-failed", "Page ExtGState page identity changed.");
      }
    }
    const inspected = await inspectMraiNativePageExtGStates(
      prepared.candidateBytes,
      pageExtGStateRequest.pageNumber,
    );
    const entries = inspected.entries.filter(
      (entry) => entry.resourceName === pageExtGStateRequest.resourceName,
    );
    if (
      entries.length !== 1 ||
      !entries[0].supportedAppearanceOnly ||
      entries[0].fillOpacity !== pageExtGStateRequest.fillOpacity ||
      entries[0].strokeOpacity !== pageExtGStateRequest.strokeOpacity ||
      entries[0].blendMode !== pageExtGStateRequest.blendMode ||
      (!pageExtGState.reused &&
        (inspected.resourceOwner.objectNumber !==
          pageExtGState.resources.objectNumber ||
          inspected.resourceOwner.generation !==
            pageExtGState.resources.generation ||
          inspected.resourceOwner.dictionaryKind !== "resource-dictionary"))
    ) {
      fail("verification-failed", "Reopened page ExtGState is inconsistent.");
    }
    const originalOwnerIsPage =
      pageExtGStateRequest.resourceOwner.objectNumber ===
        pageExtGStateRequest.page.objectNumber &&
      pageExtGStateRequest.resourceOwner.generation ===
        pageExtGStateRequest.page.generation;
    if (!originalOwnerIsPage) {
      const originalOwner = parseIndirectObject(
        prepared.candidateBytes,
        reopened,
        pageExtGStateRequest.resourceOwner.objectNumber,
        pageExtGStateRequest.resourceOwner.generation,
      );
      if (
        (await sha256PdfBytes(originalOwner.bytes)) !==
        pageExtGStateRequest.resourceOwner.expectedObjectSha256
      ) {
        fail(
          "verification-failed",
          "Original page resource owner changed during clone-on-write.",
        );
      }
    }
    if (pageExtGStateRequest.category.kind === "indirect") {
      const originalCategory = parseIndirectObject(
        prepared.candidateBytes,
        reopened,
        pageExtGStateRequest.category.target.objectNumber,
        pageExtGStateRequest.category.target.generation,
      );
      if (
        (await sha256PdfBytes(originalCategory.bytes)) !==
        pageExtGStateRequest.category.target.expectedObjectSha256
      ) {
        fail(
          "verification-failed",
          "Original ExtGState category changed during clone-on-write.",
        );
      }
    }
  }
  const pageFont = prepared.pageFontExtensionMutation;
  const pageFontRequest = prepared.request.pageFontExtensionMutation;
  if (Boolean(pageFont) !== Boolean(pageFontRequest)) {
    fail("verification-failed", "Page Font extension receipt is inconsistent.");
  }
  if (pageFont && pageFontRequest) {
    if (
      pageFont.resourceName !== pageFontRequest.fontPackage.resourceName ||
      pageFont.font.objectNumber !==
        pageFontRequest.fontPackage.font.objectNumber ||
      pageFont.font.generation !== 0 ||
      pageFont.fontObjects.length !==
        pageFontPackageObjects(pageFontRequest).length ||
      // A form-installed font's receipt names the edited Form: the shared
      // form's clone, or the unshared form itself.
      (pageFontRequest.form
        ? isolation
          ? pageFont.page.objectNumber !== isolation.primaryClone.objectNumber ||
            pageFont.page.generation !== isolation.primaryClone.generation
          : pageFont.page.objectNumber !== prepared.request.identity.objectNumber ||
            pageFont.page.generation !== prepared.request.identity.generation
        : pageFont.page.objectNumber !== pageFontRequest.page.objectNumber ||
          pageFont.page.generation !== pageFontRequest.page.generation)
    ) {
      fail("verification-failed", "Page Font extension identity changed.");
    }
    for (const item of [
      ...pageFont.fontObjects,
      pageFont.resources,
      pageFont.page,
    ]) {
      const entry = reopened.current.entries.get(item.objectNumber);
      if (
        !entry?.inUse ||
        entry.generation !== item.generation ||
        entry.offset !== item.offset
      ) {
        fail(
          "verification-failed",
          "Page Font incremental ownership is invalid.",
        );
      }
      const object = parseIndirectObject(
        prepared.candidateBytes,
        reopened,
        item.objectNumber,
        item.generation,
      );
      if ((await sha256PdfBytes(object.bytes)) !== item.sha256) {
        fail(
          "verification-failed",
          "Page Font object failed deterministic reopen verification.",
        );
      }
    }
    for (const [
      index,
      object,
    ] of pageFontRequest.fontPackage.objects.entries()) {
      const receipt = pageFont.fontObjects[index];
      if (
        receipt.objectNumber !== object.objectNumber ||
        receipt.generation !== object.generation ||
        receipt.sha256 !== object.sha256
      ) {
        fail("verification-failed", "Page Font package receipt changed.");
      }
    }
    if (pageFontRequest.form) {
      const formGraph = await resolveFormFontGraph(
        prepared.candidateBytes,
        reopened,
        pageFont.page,
      );
      if (
        !formGraph.resourcesIndirect ||
        formGraph.resourceOwnerObject.objectNumber !==
          pageFont.resources.objectNumber ||
        formGraph.resourceOwnerObject.generation !==
          pageFont.resources.generation
      ) {
        fail("verification-failed", "Reopened Form Font owner is inconsistent.");
      }
    } else {
      const pages = await inspectMraiNativeDocumentPageGraphs(
        prepared.candidateBytes,
      );
      const page = pages[pageFontRequest.pageNumber - 1];
      if (
        !page ||
        page.pageNumber !== pageFontRequest.pageNumber ||
        page.resourceAuthority.owner.objectNumber !==
          pageFont.resources.objectNumber ||
        page.resourceAuthority.owner.generation !==
          pageFont.resources.generation ||
        page.resourceAuthority.dictionaryKind !== "resource-dictionary"
      ) {
        fail("verification-failed", "Reopened page Font owner is inconsistent.");
      }
    }
    const resourceObject = parseIndirectObject(
      prepared.candidateBytes,
      reopened,
      pageFont.resources.objectNumber,
      pageFont.resources.generation,
    );
    const mappings = parseMraiNativeDirectFontDictionaryEntries(
      resourceObject.bytes,
      "resource-dictionary",
      "Reopened page",
    ).filter(
      (entry) =>
        entry.resourceName === pageFontRequest.fontPackage.resourceName,
    );
    if (
      mappings.length !== 1 ||
      mappings[0].target.objectNumber !== pageFont.font.objectNumber ||
      mappings[0].target.generation !== pageFont.font.generation
    ) {
      fail(
        "verification-failed",
        "Reopened page Font mapping is inconsistent.",
      );
    }
    const originalOwnerIsPage =
      pageFontRequest.resourceOwner.objectNumber ===
        pageFontRequest.page.objectNumber &&
      pageFontRequest.resourceOwner.generation ===
        pageFontRequest.page.generation;
    if (!originalOwnerIsPage) {
      const originalOwner = parseIndirectObject(
        prepared.candidateBytes,
        reopened,
        pageFontRequest.resourceOwner.objectNumber,
        pageFontRequest.resourceOwner.generation,
      );
      if (
        (await sha256PdfBytes(originalOwner.bytes)) !==
        pageFontRequest.resourceOwner.expectedObjectSha256
      ) {
        fail(
          "verification-failed",
          "Original page Font owner changed during clone-on-write.",
        );
      }
    }
    if (pageFontRequest.category.kind === "indirect") {
      const originalCategory = parseIndirectObject(
        prepared.candidateBytes,
        reopened,
        pageFontRequest.category.target.objectNumber,
        pageFontRequest.category.target.generation,
      );
      if (
        (await sha256PdfBytes(originalCategory.bytes)) !==
        pageFontRequest.category.target.expectedObjectSha256
      ) {
        fail(
          "verification-failed",
          "Original Font category changed during clone-on-write.",
        );
      }
    }
  }
  const fontReference = prepared.pageFontReference;
  const fontReferenceRequest = prepared.request.pageFontReference;
  if (Boolean(fontReference) !== Boolean(fontReferenceRequest)) {
    fail("verification-failed", "Page Font reference receipt is inconsistent.");
  }
  if (fontReference && fontReferenceRequest) {
    // What the source gives this reference, derived again from the source:
    // the reopened candidate holds exactly those two objects, the page's
    // Resources with these entries more and nothing else another, and the
    // page naming them and nothing else of it another.
    const expected = await resolvePageFontReference(
      sourceBytes,
      parsePdf(sourceBytes, false),
      fontReferenceRequest.pageNumber,
      fontReferenceRequest.fonts,
      prepared.request.identity,
    );
    if (
      pageFontReferenceSummary(expected.plan) !== pageFontReferenceSummary(fontReferenceRequest) ||
      JSON.stringify(fontReference.fonts) !== JSON.stringify(fontReferenceRequest.fonts.map((item) => ({
        resourceName: item.resourceName,
        font: { objectNumber: item.font.objectNumber, generation: item.font.generation },
      }))) ||
      fontReference.resources.objectNumber !== expected.plan.resources.objectNumber ||
      fontReference.resources.generation !== 0 ||
      fontReference.resources.sha256 !== (await sha256PdfBytes(expected.resourcesObject)) ||
      fontReference.page.objectNumber !== fontReferenceRequest.page.objectNumber ||
      fontReference.page.generation !== fontReferenceRequest.page.generation ||
      fontReference.page.sha256 !== (await sha256PdfBytes(expected.revisedPage))
    ) {
      fail("verification-failed", "Page Font reference identity changed.");
    }
    for (const item of [fontReference.resources, fontReference.page]) {
      const entry = reopened.current.entries.get(item.objectNumber);
      if (!entry?.inUse || entry.generation !== item.generation || entry.offset !== item.offset) {
        fail("verification-failed", "Page Font reference incremental ownership is invalid.");
      }
      const object = parseIndirectObject(prepared.candidateBytes, reopened, item.objectNumber, item.generation);
      if ((await sha256PdfBytes(object.bytes)) !== item.sha256) {
        fail("verification-failed", "Page Font reference object failed deterministic reopen verification.");
      }
    }
    const pages = await inspectMraiNativeDocumentPageGraphs(prepared.candidateBytes);
    const page = pages[fontReferenceRequest.pageNumber - 1];
    if (
      !page ||
      page.pageNumber !== fontReferenceRequest.pageNumber ||
      page.resourceAuthority.owner.objectNumber !== fontReference.resources.objectNumber ||
      page.resourceAuthority.owner.generation !== 0 ||
      page.resourceAuthority.dictionaryKind !== "resource-dictionary"
    ) {
      fail("verification-failed", "Reopened page Font owner is inconsistent.");
    }
    const mappings = parseMraiNativeDirectFontDictionaryEntries(
      parseIndirectObject(prepared.candidateBytes, reopened, fontReference.resources.objectNumber, 0).bytes,
      "resource-dictionary",
      "Reopened page",
    );
    for (const item of fontReferenceRequest.fonts) {
      const named = mappings.filter((entry) => entry.resourceName === item.resourceName);
      if (
        named.length !== 1 || named[0].inlineDictionary !== undefined ||
        named[0].target.objectNumber !== item.font.objectNumber ||
        named[0].target.generation !== item.font.generation
      ) {
        fail("verification-failed", "Reopened page Font mapping is inconsistent.");
      }
    }
    // Nothing the reference reads was written: each font, the dictionary the
    // page's resources came from and its Font category stand as they stood.
    const stands = [
      ...fontReferenceRequest.fonts.map((item) => item.font),
      ...(sameNativeObjectRef(fontReferenceRequest.resourceOwner, fontReferenceRequest.page)
        ? [] : [fontReferenceRequest.resourceOwner]),
      ...(fontReferenceRequest.category.kind === "indirect" ? [fontReferenceRequest.category.target] : []),
    ];
    for (const item of stands) {
      const object = parseIndirectObject(prepared.candidateBytes, reopened, item.objectNumber, item.generation);
      if (
        reopened.current.entries.has(item.objectNumber) ||
        (await sha256PdfBytes(object.bytes)) !== item.expectedObjectSha256
      ) {
        fail("verification-failed", "An object a page Font reference only reads was written.");
      }
    }
  }
  const appendReceipt = prepared.contentAppend;
  const appendRequest = prepared.request.contentAppend;
  if (Boolean(appendReceipt) !== Boolean(appendRequest)) {
    fail("verification-failed", "Content append receipt is inconsistent.");
  }
  if (appendReceipt && appendRequest) {
    const receipts = [
      ...appendReceipt.objects,
      ...(appendReceipt.resources ? [appendReceipt.resources] : []),
      ...(appendReceipt.page ? [appendReceipt.page] : []),
    ];
    for (const item of receipts) {
      const entry = reopened.current.entries.get(item.objectNumber);
      if (
        !entry?.inUse ||
        entry.generation !== item.generation ||
        entry.offset !== item.offset
      ) {
        fail(
          "verification-failed",
          "Content append incremental ownership is invalid.",
        );
      }
      const object = parseIndirectObject(
        prepared.candidateBytes,
        reopened,
        item.objectNumber,
        item.generation,
      );
      if ((await sha256PdfBytes(object.bytes)) !== item.sha256) {
        fail(
          "verification-failed",
          "Content append object failed deterministic reopen verification.",
        );
      }
    }
    if (appendRequest.additions.length === 0) {
      if (receipts.length !== 0) {
        fail("verification-failed", "Content append receipt has stray objects.");
      }
    } else {
      if (!appendReceipt.resources || !appendReceipt.page) {
        fail("verification-failed", "Content append Resources receipt is missing.");
      }
      const pages = await inspectMraiNativeDocumentPageGraphs(
        prepared.candidateBytes,
      );
      const page = pages[appendRequest.pageNumber - 1];
      if (
        !page ||
        page.pageNumber !== appendRequest.pageNumber ||
        !sameNativeObjectRef(page.page, appendRequest.page) ||
        !sameNativeObjectRef(
          page.resourceAuthority.owner,
          appendReceipt.resources,
        ) ||
        page.resourceAuthority.dictionaryKind !== "resource-dictionary"
      ) {
        fail(
          "verification-failed",
          "Reopened content append Resources owner is inconsistent.",
        );
      }
      for (const addition of appendRequest.additions) {
        const graph = await resolvePageResourceCategoryGraph(
          prepared.candidateBytes,
          reopened,
          page,
          addition.category,
        );
        const reference =
          addition.category === "Font"
            ? addition.fontPackage.font
            : addition.reference;
        const mappings = graph.entries.filter(
          (entry) => entry.resourceName === addition.resourceName,
        );
        if (
          mappings.length !== 1 ||
          !sameNativeObjectRef(mappings[0].target, reference)
        ) {
          fail(
            "verification-failed",
            `Reopened content append ${addition.category} mapping is inconsistent.`,
          );
        }
      }
      const originalOwnerIsPage = sameNativeObjectRef(
        appendRequest.resourceOwner,
        appendRequest.page,
      );
      if (!originalOwnerIsPage) {
        const originalOwner = parseIndirectObject(
          prepared.candidateBytes,
          reopened,
          appendRequest.resourceOwner.objectNumber,
          appendRequest.resourceOwner.generation,
        );
        if (
          (await sha256PdfBytes(originalOwner.bytes)) !==
          appendRequest.resourceOwner.expectedObjectSha256
        ) {
          fail(
            "verification-failed",
            "Original content append owner changed during clone-on-write.",
          );
        }
      }
      for (const state of Object.values(appendRequest.categories)) {
        if (!state || state.kind !== "indirect") continue;
        const originalCategory = parseIndirectObject(
          prepared.candidateBytes,
          reopened,
          state.target.objectNumber,
          state.target.generation,
        );
        if (
          (await sha256PdfBytes(originalCategory.bytes)) !==
          state.target.expectedObjectSha256
        ) {
          fail(
            "verification-failed",
            "Original content append category changed during clone-on-write.",
          );
        }
      }
    }
  }
  const consolidation = prepared.pageContentConsolidation;
  const consolidationRequest = prepared.request.pageContentConsolidation;
  if (Boolean(consolidation) !== Boolean(consolidationRequest)) {
    fail("verification-failed", "Page content consolidation receipt is inconsistent.");
  }
  if (consolidation && consolidationRequest) {
    if (
      consolidation.stream.objectNumber !== consolidationRequest.stream.objectNumber ||
      consolidation.stream.generation !== 0 ||
      consolidation.page.objectNumber !== consolidationRequest.page.objectNumber ||
      consolidation.page.generation !== consolidationRequest.page.generation
    ) {
      fail("verification-failed", "Page content consolidation identity changed.");
    }
    for (const item of [consolidation.stream, consolidation.page]) {
      const entry = reopened.current.entries.get(item.objectNumber);
      if (!entry?.inUse || entry.generation !== item.generation || entry.offset !== item.offset) {
        fail("verification-failed", "Page content consolidation incremental ownership is invalid.");
      }
      const object = parseIndirectObject(prepared.candidateBytes, reopened, item.objectNumber, item.generation);
      if ((await sha256PdfBytes(object.bytes)) !== item.sha256) {
        fail("verification-failed", "Page content consolidation object failed deterministic reopen verification.");
      }
    }
    // The reopened page draws exactly the one new stream.
    const pages = await inspectMraiNativeDocumentPageGraphs(prepared.candidateBytes);
    const page = pages[consolidationRequest.pageNumber - 1];
    if (
      !page ||
      page.contentStreams.length !== 1 ||
      page.contentStreams[0].objectNumber !== consolidationRequest.stream.objectNumber ||
      page.contentStreams[0].generation !== 0
    ) {
      fail("verification-failed", "Reopened page does not draw the consolidated stream.");
    }
  }
  const changedIdentity = isolation?.primaryClone ?? prepared.request.identity;
  const changedObject = parseIndirectObject(
    prepared.candidateBytes,
    reopened,
    changedIdentity.objectNumber,
    changedIdentity.generation,
  );
  assertTargetStream(changedObject);
  if (
    (await sha256PdfBytes(changedObject.bytes)) !==
      prepared.changedObjectSha256 ||
    (await sha256PdfBytes(changedObject.encodedStreamBytes)) !==
      prepared.changedStreamSha256 ||
    (await sha256PdfBytes(changedObject.streamBytes)) !==
      prepared.changedDecodedStreamSha256
  ) {
    fail(
      "verification-failed",
      "Reopened content stream is not deterministic.",
    );
  }
  await assertUntouchedPreconditions(
    prepared.candidateBytes,
    reopened,
    (prepared.request.untouchedObjects ?? []).filter(
      (item) =>
        (!pageFont ||
          ![...pageFont.fontObjects, pageFont.resources, pageFont.page].some(
            (revision) =>
              revision.objectNumber === item.objectNumber &&
              revision.generation === item.generation,
          )) &&
        (!fontReference ||
          !(fontReference.page.objectNumber === item.objectNumber &&
            fontReference.page.generation === item.generation)) &&
        (!consolidation ||
          !(consolidation.page.objectNumber === item.objectNumber &&
            consolidation.page.generation === item.generation)),
    ),
  );
  // Each structure element stands where the revision wrote it and reads
  // the text its revision names.
  const structRequest = prepared.request.structActualText ?? [];
  const structPrepared = prepared.structActualText ?? [];
  if (structRequest.length !== structPrepared.length) {
    fail("verification-failed", "Structure /ActualText receipt is inconsistent.");
  }
  for (const [index, item] of structPrepared.entries()) {
    const revision = structRequest[index];
    const entry = reopened.current.entries.get(item.objectNumber);
    if (
      revision.objectNumber !== item.objectNumber || revision.generation !== item.generation ||
      !entry?.inUse || entry.kind !== "direct" || entry.generation !== item.generation || entry.offset !== item.offset
    ) {
      fail("verification-failed", "Structure element incremental ownership is invalid.");
    }
    const reopenedElement = parseIndirectObject(
      prepared.candidateBytes, reopened, item.objectNumber, item.generation, undefined, "opaque",
    );
    const written = readMraiNativeStructElementActualText(reopenedElement.bytes);
    const writtenBytes = written ? decodeMraiNativePdfStringToken(written.token) : null;
    if (
      (await sha256PdfBytes(reopenedElement.bytes)) !== item.sha256 ||
      !written || !bytesEqual(written.token, revision.value) || !writtenBytes ||
      decodeMraiNativePdfTextString(writtenBytes, false) !== revision.text
    ) {
      fail("verification-failed", "Reopened structure element does not read its new /ActualText.");
    }
  }
  // The catalog's packet stands where the revision wrote it and claims no
  // PDF/A conformance any more.
  const pdfa = prepared.pdfaIdentificationRemoval;
  if (pdfa) {
    const entry = reopened.current.entries.get(pdfa.objectNumber);
    if (!entry?.inUse || entry.kind !== "direct" || entry.generation !== pdfa.generation || entry.offset !== pdfa.offset) {
      fail("verification-failed", "The catalog's Metadata incremental ownership is invalid.");
    }
    const reopenedPacket = parseIndirectObject(
      prepared.candidateBytes, reopened, pdfa.objectNumber, pdfa.generation, undefined, "opaque",
    );
    const identification = reopenedPacket.streamBytes
      ? readMraiNativePdfaIdentificationRanges(asLatin1(reopenedPacket.streamBytes))
      : null;
    if ((await sha256PdfBytes(reopenedPacket.bytes)) !== pdfa.sha256 || !identification || identification.claims) {
      fail("verification-failed", "The reopened catalog Metadata still claims PDF/A.");
    }
  }
  // The catalog stating a later version stands where the revision wrote it.
  const raised = prepared.catalogVersionRaise;
  if (raised) {
    const entry = reopened.current.entries.get(raised.objectNumber);
    if (!entry?.inUse || entry.kind !== "direct" || entry.generation !== raised.generation || entry.offset !== raised.offset) {
      fail("verification-failed", "The catalog's incremental ownership is invalid.");
    }
    const reopenedCatalog = parseIndirectObject(
      prepared.candidateBytes, reopened, raised.objectNumber, raised.generation, undefined, "opaque",
    );
    if ((await sha256PdfBytes(reopenedCatalog.bytes)) !== raised.sha256) {
      fail("verification-failed", "The reopened catalog is not the one written.");
    }
  }
  return Object.freeze({
    verified: true,
    identity: prepared.request.identity,
    sourceXrefOffset: prepared.sourceXrefOffset,
    incrementalXrefOffset: prepared.incrementalXrefOffset,
    candidateSha256: prepared.candidateSha256,
  });
}

// Token-level readers shared by source-bound font adapters.
export { oneTopLevelReference as readMraiNativeDictionaryReference };

export { parseDirectXObjectDictionaryEntries as readMraiNativeDirectXObjectEntries };

/** An indirect object's body: what stands between its `n g obj` and its
 * `endobj` (a compressed member is all body), without the white space
 * round it. */
function indirectObjectBody(object: ParsedIndirectObject): Uint8Array {
  let start = 0;
  let end = object.bytes.length;
  if (object.storage === "direct") {
    const text = asLatin1(object.bytes);
    const header = /^(\d+)\s+(\d+)\s+obj\b/.exec(text);
    const close = text.lastIndexOf("endobj");
    if (!header || close < header[0].length) {
      fail("unsupported-document", "PDF indirect object is malformed.");
    }
    start = header[0].length;
    end = close;
  }
  const white = (byte: number) => byte === 0x00 || byte === 0x09 || byte === 0x0a || byte === 0x0c ||
    byte === 0x0d || byte === 0x20;
  while (start < end && white(object.bytes[start])) start += 1;
  while (end > start && white(object.bytes[end - 1])) end -= 1;
  return object.bytes.subarray(start, end);
}

/**
 * Undo or Redo of one revision the engine appended, planned on a later head
 * of the same chain (engine unification, stage 2: the history of an edit
 * published as its bytes). Every object the revision wrote is written again
 * as it stood before it (undo) or as the revision wrote it (redo). `head`
 * holds the revision: an incremental chain only appends, so its first
 * `afterLength` bytes are the file the revision made and its first
 * `beforeLength` the file it was appended to. Each object must still stand
 * in `head` as the other direction left it: an object a later edit changed
 * makes the step unavailable rather than overwriting that edit. An object the
 * revision created stays where it is; nothing the restore writes refers to it
 * on undo, and it is the revision's own again on redo.
 */
/**
 * A dictionary object a later edit changed too, put back key by key: the
 * keys the step changed go back (each must still hold what the step left),
 * every other key keeps what the file holds now. A page's dictionary is the
 * case: a content edit changes its /Contents, a note written since its
 * /Annots, and the content edit's Undo or Redo takes back its /Contents
 * alone. Null for a stream, a value that is not a plain dictionary, a key
 * stated twice, or a key both the step and a later edit changed.
 */
function mergedRestoreDictionary(
  expected: Uint8Array,
  desired: Uint8Array,
  standing: Uint8Array,
): Uint8Array | null {
  const read = (bytes: Uint8Array) => {
    const text = asLatin1(bytes).trim();
    if (!text.startsWith("<<") || !text.endsWith(">>") || /\bstream\b/.test(text.slice(text.lastIndexOf(">>")))) return null;
    try {
      if (dictionaryEnd(text, 0) !== text.length) return null;
      const entries = topLevelDictionaryEntries(text);
      const values = new Map<string, Readonly<{ name: string; value: string }>>();
      for (const entry of entries) {
        if (values.has(entry.key)) return null;
        values.set(entry.key, Object.freeze({ name: text.slice(entry.keyStart, entry.valueStart).trim(),
          value: text.slice(entry.valueStart, entry.valueEnd).trim() }));
      }
      return values;
    } catch {
      return null;
    }
  };
  const was = read(expected);
  const goal = read(desired);
  const now = read(standing);
  if (!was || !goal || !now) return null;
  const result = new Map(now);
  for (const key of new Set([...was.keys(), ...goal.keys()])) {
    const from = was.get(key)?.value;
    const to = goal.get(key);
    if (from === to?.value) continue;
    if (now.get(key)?.value !== from) return null;
    if (to) result.set(key, to);
    else result.delete(key);
  }
  return latin1Bytes(`<< ${[...result.values()].map((entry) => `${entry.name} ${entry.value}`).join(" ")} >>`);
}

/** The keys of a stream dictionary that say how its bytes are stored, not
 * what they are. */
const STREAM_STORAGE_KEYS = new Set(["Length", "Filter", "DecodeParms", "DL"]);

/**
 * Whether two objects are one stream stored another way: the same decoded
 * content and the same dictionary but for its storage keys. A version the
 * server wrote whole (a save its incremental writer could not take:
 * `pdf_incremental_patchback mode=full`) stores every stream again -- other
 * compression, another length -- and a step of the working copy whose object
 * stands so read as changed by a later edit and could not be undone
 * (pdf-walk-1000 rp1014: a line's typing undone after the Undo of a box made
 * of another line was written whole by the server).
 */
function sameStreamStoredOtherwise(left: ParsedIndirectObject, right: ParsedIndirectObject): boolean {
  if (!left.streamBytes || !right.streamBytes || !bytesEqual(left.streamBytes, right.streamBytes)) return false;
  const kept = (object: ParsedIndirectObject) => {
    const text = asLatin1(object.bytes);
    return topLevelDictionaryEntries(text).filter((entry) => !STREAM_STORAGE_KEYS.has(entry.key))
      .map((entry) => `${entry.key} ${text.slice(entry.valueStart, entry.valueEnd).replace(/\s+/g, " ").trim()}`)
      .sort().join("\n");
  };
  return kept(left) === kept(right);
}

export function planMraiNativeObjectRestore(input: Readonly<{
  headBytes: Uint8Array;
  beforeLength: number;
  afterLength: number;
  direction: "undo" | "redo";
}>): Readonly<{ candidateBytes: Uint8Array; restoredObjects: number }> {
  const { headBytes, beforeLength, afterLength, direction } = input;
  if (
    !Number.isSafeInteger(beforeLength) ||
    !Number.isSafeInteger(afterLength) ||
    beforeLength < 1 ||
    afterLength <= beforeLength ||
    afterLength > headBytes.length
  ) {
    fail("invalid-request", "The revision to restore is not in this file.");
  }
  const before = headBytes.subarray(0, beforeLength);
  const after = headBytes.subarray(0, afterLength);
  const beforeParsed = parsePdf(before, false);
  const afterParsed = parsePdf(after, false);
  const headParsed = parsePdf(headBytes, false);
  const revision = afterParsed.current;
  if (afterParsed.sections.length !== beforeParsed.sections.length + 1 || revision.offset < beforeLength) {
    fail("invalid-request", "The revision to restore is not one appended revision.");
  }
  const beforeLatest = latestEntries(beforeParsed);
  const headLatest = latestEntries(headParsed);
  const read = (bytes: Uint8Array, parsed: ParsedPdf, objectNumber: number, generation: number) =>
    parseIndirectObject(bytes, parsed, objectNumber, generation, undefined, "raw");
  const revisions: IncrementalObjectRevision[] = [];
  for (const [objectNumber, entry] of [...revision.entries].sort(([left], [right]) => left - right)) {
    // Rows the revision only stated again (an object the chain had to be read
    // round, where it really is) and its own cross-reference stream are no
    // object it wrote.
    if (!entry.inUse || entry.offset < beforeLength ||
      (revision.kind === "stream" && entry.kind === "direct" && entry.offset === revision.offset)) continue;
    const written = read(after, afterParsed, objectNumber, entry.generation);
    const previous = beforeLatest.get(objectNumber);
    const existed = previous !== undefined && previous.inUse;
    if (!existed) continue;
    const prior = read(before, beforeParsed, objectNumber, previous.generation);
    const current = headLatest.get(objectNumber);
    if (!current?.inUse) {
      fail("invalid-request", "An object this step restores is no longer in the file.");
    }
    const standing = indirectObjectBody(read(headBytes, headParsed, objectNumber, current.generation));
    const [expected, desired] = direction === "undo" ? [written, prior] : [prior, written];
    let body = indirectObjectBody(desired);
    if (!bytesEqual(standing, indirectObjectBody(expected))) {
      const storedOtherwise = (() => {
        try {
          const expectedDecoded = direction === "undo"
            ? parseIndirectObject(after, afterParsed, objectNumber, entry.generation, undefined, "decoded")
            : parseIndirectObject(before, beforeParsed, objectNumber, previous.generation, undefined, "decoded");
          return sameStreamStoredOtherwise(expectedDecoded,
            parseIndirectObject(headBytes, headParsed, objectNumber, current.generation, undefined, "decoded"));
        } catch {
          return false;
        }
      })();
      if (!storedOtherwise) {
        const merged = mergedRestoreDictionary(indirectObjectBody(expected), body, standing);
        if (!merged) fail("invalid-request", "A later edit changed what this step would restore.");
        body = merged;
      }
    }
    revisions.push(Object.freeze({
      objectNumber,
      generation: desired.generation,
      bytes: concatBytes([
        ascii(`${objectNumber} ${desired.generation} obj\n`),
        body,
        ascii("\nendobj"),
      ]),
    }));
  }
  if (!revisions.length) fail("invalid-request", "This step writes nothing back.");
  const written = makeMultiObjectIncrementalRevision(headBytes, headParsed.current, revisions, {
    min: 1,
    max: 100_000,
    label: "Object restore",
  });
  return Object.freeze({ candidateBytes: written.candidateBytes, restoredObjects: revisions.length });
}

/** The decompressed size one kept body may name (the server's bound on a
 * restore evidence field, native_surgical_contract._restore_bytes). */
const MAX_REVISION_RESTORE_BODY_BYTES = 4 * 1024 * 1024;

/** A body a revision's restore evidence keeps: base64 of a zlib stream (the
 * server's pack_native_restore_bytes), read within its bound. */
function unpackRevisionRestoreBody(value: string): Uint8Array {
  let packed: Uint8Array | null = null;
  if (/^[A-Za-z0-9+/]+={0,2}$/.test(value) && value.length % 4 === 0) {
    try {
      packed = Uint8Array.from(atob(value), (character) => character.charCodeAt(0));
    } catch {
      packed = null;
    }
  }
  if (!packed) fail("invalid-request", "The kept revision of this step is unreadable.");
  const chunks: Uint8Array[] = [];
  let length = 0;
  let complete = false;
  try {
    const inflater = new Unzlib((chunk, final) => {
      length += chunk.length;
      if (length > MAX_REVISION_RESTORE_BODY_BYTES) throw new Error("kept body exceeds its bound");
      if (chunk.length) chunks.push(Uint8Array.from(chunk));
      if (final) complete = true;
    });
    inflater.push(packed, true);
  } catch {
    complete = false;
  }
  if (!complete) fail("invalid-request", "The kept revision of this step is unreadable.");
  return concatBytes(chunks);
}

/**
 * Undo or Redo of a revision the engine appended, once a source consolidation
 * has dropped its two sides (the chain was rewritten as one section): planned
 * from the restore evidence the consolidation record kept for it
 * (readPdfNativeRevisionRestoreEvidence) instead of from prefixes of `head`.
 * The rule is planMraiNativeObjectRestore's: every object the revision wrote
 * over one that already stood must still stand in `head` as the other
 * direction left it -- its body (between `N G obj` and `endobj`, without the
 * white space round it) equal to the evidence's `after` for an undo, to its
 * `before` for a redo; an object a later edit changed makes the step
 * unavailable rather than overwriting that edit -- and is written again, in
 * one appended revision, with the other body under the generation that side
 * wrote it under. An object the revision created stays where it is in both
 * directions (the evidence names it by digest only). A consolidation keeps
 * every object's bytes under its own number and generation, so an object no
 * later edit changed holds its `after` body in every later head.
 */
export function planMraiNativeObjectRestoreFromEvidence(input: Readonly<{
  headBytes: Uint8Array;
  evidence: PdfNativeRevisionRestoreEvidence;
  direction: "undo" | "redo";
}>): Readonly<{ candidateBytes: Uint8Array; restoredObjects: number }> {
  const { headBytes, evidence, direction } = input;
  if ((direction !== "undo" && direction !== "redo") || !isPdfNativeRevisionRestoreEvidence(evidence)) {
    fail("invalid-request", "The kept revision of this step is unreadable.");
  }
  const headParsed = parsePdf(headBytes, false);
  const headLatest = latestEntries(headParsed);
  const revisions: IncrementalObjectRevision[] = [];
  for (const object of evidence.objects) {
    const after = unpackRevisionRestoreBody(object.after);
    const before = unpackRevisionRestoreBody(object.before);
    const current = headLatest.get(object.objectNumber);
    if (!current?.inUse) {
      fail("invalid-request", "An object this step restores is no longer in the file.");
    }
    const standing = indirectObjectBody(
      parseIndirectObject(headBytes, headParsed, object.objectNumber, current.generation, undefined, "raw"),
    );
    const [expected, desired, generation] = direction === "undo"
      ? [after, before, object.beforeGeneration]
      : [before, after, object.generation];
    let body = desired;
    if (!bytesEqual(standing, expected)) {
      const merged = mergedRestoreDictionary(expected, desired, standing);
      if (!merged) fail("invalid-request", "A later edit changed what this step would restore.");
      body = merged;
    }
    revisions.push(Object.freeze({
      objectNumber: object.objectNumber,
      generation,
      bytes: concatBytes([ascii(`${object.objectNumber} ${generation} obj\n`), body, ascii("\nendobj")]),
    }));
  }
  if (!revisions.length) fail("invalid-request", "This step writes nothing back.");
  const written = makeMultiObjectIncrementalRevision(headBytes, headParsed.current, revisions, {
    min: 1,
    max: 100_000,
    label: "Object restore",
  });
  return Object.freeze({ candidateBytes: written.candidateBytes, restoredObjects: revisions.length });
}

/**
 * The marked-content sequence a picture's paint is drawn in, read from the
 * stream that draws it: the innermost `BDC` before the paint that names an
 * MCID and is still open there (ISO 32000-1 §14.6, §14.7.4). A tagged
 * document's structure element for the picture (a /Figure) reaches it by
 * that MCID on the picture's page. Null for a paint inside a form, or one no
 * numbered sequence holds (an untagged page).
 */
export async function readMraiNativeEmbeddedImageMarkedContent(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly imageOrder: number;
}): Promise<Readonly<{ mcid: number; page: Readonly<{ objectNumber: number; generation: number }> }> | null> {
  const target = await resolveEmbeddedImageTransformTarget({
    sourceBytes: input.sourceBytes,
    pageNumber: exactPositiveInteger(input.pageNumber, "embedded image page number"),
    imageOrder: exactPositiveInteger(input.imageOrder, "embedded image order", true),
  });
  if (target.invocation.host) return null;
  const text = decodeLatin1(target.inspected.streamBytes.subarray(0, target.span.start));
  const open: (number | null)[] = [];
  const operands: string[] = [];
  let index = 0;
  while (index < text.length) {
    const char = text[index];
    if (/\s/.test(char)) { index += 1; continue; }
    if (char === "%") {
      while (index < text.length && text[index] !== "\n" && text[index] !== "\r") index += 1;
      continue;
    }
    if (char === "(") {
      // A string: skipped whole (escapes and balanced parentheses).
      let depth = 0;
      for (; index < text.length; index += 1) {
        if (text[index] === "\\") { index += 1; continue; }
        if (text[index] === "(") depth += 1;
        else if (text[index] === ")" && --depth === 0) { index += 1; break; }
      }
      operands.push("()");
      continue;
    }
    if (char === "<" && text[index + 1] === "<") {
      // A dictionary operand (a BDC property list): kept whole.
      let depth = 0;
      const start = index;
      for (; index < text.length; index += 1) {
        if (text[index] === "(") {
          let strDepth = 0;
          for (; index < text.length; index += 1) {
            if (text[index] === "\\") { index += 1; continue; }
            if (text[index] === "(") strDepth += 1;
            else if (text[index] === ")" && --strDepth === 0) break;
          }
          continue;
        }
        if (text[index] === "<" && text[index + 1] === "<") { depth += 1; index += 1; continue; }
        if (text[index] === ">" && text[index + 1] === ">") {
          depth -= 1; index += 1;
          if (depth === 0) { index += 1; break; }
        }
      }
      operands.push(text.slice(start, index));
      continue;
    }
    if (char === "<") {
      const end = text.indexOf(">", index);
      index = end < 0 ? text.length : end + 1;
      operands.push("<>");
      continue;
    }
    if (char === "[" || char === "]") { index += 1; continue; }
    let end = index + 1;
    while (end < text.length && !/[\s()<>[\]{}/%]/.test(text[end])) end += 1;
    const token = text.slice(index, end);
    index = end;
    if (token === "BI") {
      // Inline image data: to its EI.
      const stop = text.indexOf("EI", index);
      index = stop < 0 ? text.length : stop + 2;
      operands.length = 0;
      continue;
    }
    if (token === "BDC") {
      const properties = operands.at(-1) ?? "";
      const mcid = /\/MCID\s+(\d+)/.exec(properties);
      open.push(properties.startsWith("<<") && mcid ? Number(mcid[1]) : null);
    } else if (token === "BMC") {
      open.push(null);
    } else if (token === "EMC") {
      open.pop();
    }
    if (token.startsWith("/") || /^[-+.\d]/.test(token) || token === "true" || token === "false" || token === "null") {
      operands.push(token);
    } else {
      operands.length = 0;
    }
  }
  for (let level = open.length - 1; level >= 0; level -= 1) {
    const mcid = open[level];
    if (mcid !== null) return Object.freeze({ mcid, page: Object.freeze({ ...target.graph.page.page }) });
  }
  return null;
}

/** A text string as a PDF writes it for any language: UTF-16BE with its byte
 * order mark, in hex (ISO 32000-1 §7.9.2.2). */
function pdfUnicodeTextStringToken(text: string): Uint8Array {
  let hex = "<FEFF";
  for (let index = 0; index < text.length; index += 1) {
    hex += text.charCodeAt(index).toString(16).padStart(4, "0").toUpperCase();
  }
  return ascii(`${hex}>`);
}

/** The longest alternate description written (Word's own limit is far
 * smaller; this only keeps a revision bounded). */
export const MRAI_NATIVE_ALT_TEXT_MAX_LENGTH = 4000;

/**
 * One structure element's alternate description (/Alt, ISO 32000-1 §14.9.3),
 * set, replaced or (empty text) taken away, written as an incremental
 * revision of that element alone: nothing else in the file changes. The
 * element must be a dictionary (it always is); one held in an object stream
 * is written again under a header of its own.
 */
export async function planMraiNativeStructElementAlt(input: {
  readonly sourceBytes: Uint8Array;
  readonly element: Readonly<{ objectNumber: number; generation: number }>;
  readonly alt: string;
}): Promise<Readonly<{ candidateBytes: Uint8Array; candidateSha256: string; previous: string | null }>> {
  assertObjectReference(input.element.objectNumber, input.element.generation);
  if (typeof input.alt !== "string" || input.alt.length > MRAI_NATIVE_ALT_TEXT_MAX_LENGTH || !isWellFormedUtf16(input.alt)) {
    fail("invalid-request", "Alternate text is invalid.");
  }
  const parsed = parseSourceStructure(input.sourceBytes);
  const object = parseIndirectObject(input.sourceBytes, parsed, input.element.objectNumber, input.element.generation);
  if (object.streamStart !== null) {
    fail("unsupported-document", "A structure element is a dictionary, not a stream.");
  }
  const source = asLatin1(object.bytes);
  const entries = topLevelDictionaryEntries(source);
  const existing = entries.filter((entry) => entry.key === "Alt");
  if (existing.length > 1) fail("unsupported-document", "The structure element names /Alt twice.");
  const previousToken = existing[0] ? object.bytes.slice(existing[0].valueStart, existing[0].valueEnd) : null;
  const previousBytes = previousToken ? decodeMraiNativePdfStringToken(previousToken) : null;
  const previous = previousBytes ? decodeMraiNativePdfTextString(previousBytes, false) : null;
  const text = input.alt.trim();
  let revised: Uint8Array;
  if (existing[0]) {
    // The key and its value are replaced (or, empty, taken away) in place.
    const keyStart = source.lastIndexOf("/Alt", existing[0].valueStart);
    if (keyStart < 0) fail("unsupported-document", "The structure element's /Alt cannot be read.");
    revised = concatBytes([
      object.bytes.subarray(0, text ? existing[0].valueStart : keyStart),
      text ? pdfUnicodeTextStringToken(text) : new Uint8Array(),
      object.bytes.subarray(existing[0].valueEnd),
    ]);
  } else {
    if (!text) fail("invalid-request", "Embedded image transform is a no-op.");
    const open = source.indexOf("<<");
    if (open < 0) fail("unsupported-document", "The structure element is not a dictionary.");
    revised = concatBytes([
      object.bytes.subarray(0, open + 2),
      ascii("/Alt "),
      pdfUnicodeTextStringToken(text),
      ascii(" "),
      object.bytes.subarray(open + 2),
    ]);
  }
  if (previous !== null && previous === text) fail("invalid-request", "Embedded image transform is a no-op.");
  const bytes = object.storage === "direct"
    ? revised
    : concatBytes([ascii(`${object.objectNumber} ${object.generation} obj\n`), revised, ascii("\nendobj")]);
  const incremental = makeMultiObjectIncrementalRevision(input.sourceBytes, parsed.current,
    [{ objectNumber: object.objectNumber, generation: object.generation, bytes }],
    { min: 1, max: 1, label: "Alternate text" });
  return Object.freeze({
    candidateBytes: incremental.candidateBytes,
    candidateSha256: await sha256PdfBytes(incremental.candidateBytes),
    previous,
  });
}

/** The most pictures of one page whose originals are kept named. */
export const MRAI_NATIVE_PICTURE_ORIGINS_MAX = 64;
const PICTURE_ORIGINS_VALUE =
  /^<<\s*\/MRAI\s*<<\s*\/LastModified\s*\([^()\\]*\)\s*\/Private\s*<<\s*\/Pictures\s*\[((?:\s*\[\s*\d+\s+\d+\s+R\s+\d+\s+\d+\s+R\s+<[0-9A-Fa-f]*>\s*\])*)\s*\]\s*>>\s*>>\s*>>$/;
const PICTURE_ORIGIN_ENTRY = /\[\s*(\d+)\s+(\d+)\s+R\s+(\d+)\s+(\d+)\s+R\s+<([0-9A-Fa-f]*)>\s*\]/g;

/**
 * Where a picture the reader's Picture Format redrew came from (Word keeps
 * the picture before its formats in the file, and Reset Picture returns to
 * it in any later session): the page's own page-piece data (/PieceInfo,
 * ISO 32000-1 §14.5) names, for each such image object, the image it was
 * drawn from and the format drawn (`data`, the reader's own). Written as an
 * incremental revision of the page object alone; the original image stays
 * in the file because it is named here. A /PieceInfo of another application
 * is never touched: refused, the picture is reset only within the session.
 */
export async function planMraiNativePictureOrigin(input: {
  readonly sourceBytes: Uint8Array;
  readonly page: Readonly<{ objectNumber: number; generation: number }>;
  readonly image: Readonly<{ objectNumber: number; generation: number }>;
  readonly original: Readonly<{ objectNumber: number; generation: number }>;
  readonly data: string;
}): Promise<Readonly<{ candidateBytes: Uint8Array; candidateSha256: string }>> {
  for (const ref of [input.page, input.image, input.original]) assertObjectReference(ref.objectNumber, ref.generation);
  if (typeof input.data !== "string" || input.data.length > 4096 || !/^[\x20-\x7e]*$/.test(input.data)) {
    fail("invalid-request", "Picture origin data is invalid.");
  }
  const parsed = parseSourceStructure(input.sourceBytes);
  const object = parseIndirectObject(input.sourceBytes, parsed, input.page.objectNumber, input.page.generation);
  if (object.streamStart !== null) fail("unsupported-document", "A page is a dictionary, not a stream.");
  const source = asLatin1(object.bytes);
  const entries = topLevelDictionaryEntries(source);
  const existing = entries.filter((entry) => entry.key === "PieceInfo");
  if (existing.length > 1) fail("unsupported-document", "The page names /PieceInfo twice.");
  const kept: string[] = [];
  if (existing[0]) {
    const value = source.slice(existing[0].valueStart, existing[0].valueEnd).trim();
    const ours = PICTURE_ORIGINS_VALUE.exec(value);
    if (!ours) fail("unsupported-document", "The page's /PieceInfo belongs to another application.");
    for (const match of ours[1].matchAll(PICTURE_ORIGIN_ENTRY)) {
      if (Number(match[1]) === input.image.objectNumber && Number(match[2]) === input.image.generation) continue;
      kept.push(match[0].replace(/\s+/g, " "));
    }
  }
  let hex = "";
  for (let index = 0; index < input.data.length; index += 1) {
    hex += input.data.charCodeAt(index).toString(16).padStart(2, "0").toUpperCase();
  }
  kept.push(`[${input.image.objectNumber} ${input.image.generation} R ${input.original.objectNumber} ` +
    `${input.original.generation} R <${hex}>]`);
  while (kept.length > MRAI_NATIVE_PICTURE_ORIGINS_MAX) kept.shift();
  const now = new Date();
  const stamp = `D:${now.getUTCFullYear()}${String(now.getUTCMonth() + 1).padStart(2, "0")}` +
    `${String(now.getUTCDate()).padStart(2, "0")}${String(now.getUTCHours()).padStart(2, "0")}` +
    `${String(now.getUTCMinutes()).padStart(2, "0")}${String(now.getUTCSeconds()).padStart(2, "0")}Z`;
  const value = ascii(`<< /MRAI << /LastModified (${stamp}) /Private << /Pictures [${kept.join(" ")}] >> >> >>`);
  let revised: Uint8Array;
  if (existing[0]) {
    revised = concatBytes([
      object.bytes.subarray(0, existing[0].valueStart),
      value,
      object.bytes.subarray(existing[0].valueEnd),
    ]);
  } else {
    const open = source.indexOf("<<");
    if (open < 0) fail("unsupported-document", "The page is not a dictionary.");
    revised = concatBytes([object.bytes.subarray(0, open + 2), ascii("/PieceInfo "), value, ascii(" "),
      object.bytes.subarray(open + 2)]);
  }
  const bytes = object.storage === "direct"
    ? revised
    : concatBytes([ascii(`${object.objectNumber} ${object.generation} obj\n`), revised, ascii("\nendobj")]);
  const incremental = makeMultiObjectIncrementalRevision(input.sourceBytes, parsed.current,
    [{ objectNumber: object.objectNumber, generation: object.generation, bytes }],
    { min: 1, max: 1, label: "Picture origin" });
  return Object.freeze({
    candidateBytes: incremental.candidateBytes,
    candidateSha256: await sha256PdfBytes(incremental.candidateBytes),
  });
}

/** The page object and the image object one image paint of a page draws
 * (null for a paint a form hosts, or one past the page's paints). */
export async function readMraiNativeEmbeddedImageObject(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly imageOrder: number;
}): Promise<Readonly<{ page: Readonly<{ objectNumber: number; generation: number }>;
  image: Readonly<{ objectNumber: number; generation: number }> }> | null> {
  const pageNumber = exactPositiveInteger(input.pageNumber, "embedded image page number");
  const graph = await resolveNativeImageRemovalGraph(input.sourceBytes, pageNumber);
  const invocation = graph.invocations[input.imageOrder];
  if (!invocation || invocation.host) return null;
  const image = graph.images.get(invocation.resourceName);
  if (!image) return null;
  return Object.freeze({
    page: Object.freeze({ objectNumber: graph.page.page.objectNumber, generation: graph.page.page.generation }),
    image: Object.freeze({ objectNumber: image.objectNumber, generation: image.generation }),
  });
}

/**
 * The page's own paints that draw one image object, by their order among the
 * page's image paints, with the digest of that object as these bytes hold it.
 * Two files of one page that each draw the object once, and hold it as the
 * same bytes, draw the same picture there whatever else of the page differs:
 * how a picture changed in two windows is found again when the other window
 * changed more of the page than that picture. `sha256` is null when the
 * object cannot be read here.
 */
export async function listMraiNativeEmbeddedImagePaintsOfObject(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly image: Readonly<{ objectNumber: number; generation: number }>;
}): Promise<Readonly<{ orders: readonly number[]; sha256: string | null }>> {
  const pageNumber = exactPositiveInteger(input.pageNumber, "embedded image page number");
  assertObjectReference(input.image.objectNumber, input.image.generation);
  const graph = await resolveNativeImageRemovalGraph(input.sourceBytes, pageNumber);
  const orders: number[] = [];
  graph.invocations.forEach((invocation, order) => {
    if (invocation.host) return;
    const image = graph.images.get(invocation.resourceName);
    if (image && image.objectNumber === input.image.objectNumber && image.generation === input.image.generation) orders.push(order);
  });
  let sha256: string | null = null;
  try {
    const parsed = parseSourceStructure(input.sourceBytes);
    const object = parseIndirectObject(input.sourceBytes, parsed, input.image.objectNumber, input.image.generation,
      undefined, "raw");
    if (object.streamStart !== null) sha256 = await sha256PdfBytes(object.bytes);
  } catch {
    sha256 = null;
  }
  return Object.freeze({ orders: Object.freeze(orders), sha256 });
}

/** What planMraiNativePictureOrigin wrote for `image` on `page`, or null. */
export async function readMraiNativePictureOrigin(input: {
  readonly sourceBytes: Uint8Array;
  readonly page: Readonly<{ objectNumber: number; generation: number }>;
  readonly image: Readonly<{ objectNumber: number; generation: number }>;
}): Promise<Readonly<{ original: Readonly<{ objectNumber: number; generation: number }>; data: string }> | null> {
  assertObjectReference(input.page.objectNumber, input.page.generation);
  const parsed = parseSourceStructure(input.sourceBytes);
  const object = parseIndirectObject(input.sourceBytes, parsed, input.page.objectNumber, input.page.generation);
  if (object.streamStart !== null) return null;
  const source = asLatin1(object.bytes);
  const entry = topLevelDictionaryEntries(source).find((candidate) => candidate.key === "PieceInfo");
  if (!entry) return null;
  const ours = PICTURE_ORIGINS_VALUE.exec(source.slice(entry.valueStart, entry.valueEnd).trim());
  if (!ours) return null;
  for (const match of ours[1].matchAll(PICTURE_ORIGIN_ENTRY)) {
    if (Number(match[1]) !== input.image.objectNumber || Number(match[2]) !== input.image.generation) continue;
    const hex = match[5];
    let data = "";
    for (let index = 0; index + 1 < hex.length; index += 2) data += String.fromCharCode(parseInt(hex.slice(index, index + 2), 16));
    return Object.freeze({
      original: Object.freeze({ objectNumber: Number(match[3]), generation: Number(match[4]) }),
      data,
    });
  }
  return null;
}

/**
 * One content stream with `replacement` in place of its decoded bytes
 * `start`..`end` (the bytes there must still hash to `expectedSpanSha256`),
 * written as an incremental revision of that stream object alone, encoded
 * the way the stream was. For an edit published as its bytes, whose Undo is
 * the object restore of its revision.
 */
export async function planMraiNativeContentStreamSpanRevision(input: {
  readonly sourceBytes: Uint8Array;
  readonly stream: Readonly<{ objectNumber: number; generation: number }>;
  readonly start: number;
  readonly end: number;
  readonly expectedSpanSha256: string;
  readonly replacement: Uint8Array;
}): Promise<Readonly<{ candidateBytes: Uint8Array; candidateSha256: string }>> {
  assertObjectReference(input.stream.objectNumber, input.stream.generation);
  const parsed = parseSourceStructure(input.sourceBytes);
  const object = parseIndirectObject(input.sourceBytes, parsed, input.stream.objectNumber, input.stream.generation);
  assertTargetStream(object);
  if (!Number.isSafeInteger(input.start) || !Number.isSafeInteger(input.end) || input.start < 0 ||
    input.end < input.start || input.end > object.streamBytes.length) {
    fail("invalid-request", "Content stream span is out of range.");
  }
  if ((await sha256PdfBytes(object.streamBytes.slice(input.start, input.end))) !== input.expectedSpanSha256) {
    fail("precondition-mismatch", "Content stream span changed.");
  }
  const revised = reviseContentStreamObject({ sourceBytes: input.sourceBytes, object, start: input.start,
    end: input.end, replacement: input.replacement });
  const incremental = makeMultiObjectIncrementalRevision(input.sourceBytes, parsed.current,
    [{ objectNumber: object.objectNumber, generation: object.generation, bytes: revised.bytes }],
    { min: 1, max: 1, label: "Content stream span" });
  return Object.freeze({
    candidateBytes: incremental.candidateBytes,
    candidateSha256: await sha256PdfBytes(incremental.candidateBytes),
  });
}
