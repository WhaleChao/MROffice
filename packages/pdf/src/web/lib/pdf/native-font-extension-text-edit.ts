/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { LATIN1_EXACT } from "./engine/native/latin1";
import {
  MRAI_NATIVE_FONT_EXTENSION_TEXT_ACTION,
  MRAI_NATIVE_FONT_EXTENSION_TEXT_HISTORY_ACTION,
  MRAI_NATIVE_FONT_EXTENSION_TEXT_HISTORY_SCHEMA_VERSION,
  MRAI_NATIVE_FONT_EXTENSION_TEXT_SCHEMA_VERSION,
  MRAI_NATIVE_FONT_REUSE_TEXT_ACTION,
  MRAI_NATIVE_FONT_REUSE_TEXT_SCHEMA_VERSION,
} from "./engine/native/vocabulary";
export {
  MRAI_NATIVE_FONT_EXTENSION_TEXT_ACTION,
  MRAI_NATIVE_FONT_EXTENSION_TEXT_HISTORY_ACTION,
  MRAI_NATIVE_FONT_EXTENSION_TEXT_HISTORY_SCHEMA_VERSION,
  MRAI_NATIVE_FONT_EXTENSION_TEXT_SCHEMA_VERSION,
  MRAI_NATIVE_FONT_REUSE_TEXT_ACTION,
  MRAI_NATIVE_FONT_REUSE_TEXT_SCHEMA_VERSION,
};
import type { PdfSharedStateRecord } from "./engine/shared-state-bridge";
import {
  pdfNativePdfaIdentificationRecordField,
  pdfNativePdfaIdentificationRevisedObjects,
  pdfNativeStructActualTextRecordField,
} from "./engine/adapters/mrai-native-record";
import { isPdfNativeObjectRestoreEvidence } from "./engine/native/byte-restore-model";
import { planMraiNativeStructActualTextHistory, restoreBytes } from "./engine/native/byte-restore-history";
import {
  inspectMraiNativeContentStream,
  inspectMraiNativeConsolidatedContentStream,
  inspectMraiNativeObject,
  inspectMraiNativePageContentGraph,
  inspectMraiNativeReusablePageFontExtension,
  appendMraiNativeStructActualTextRevisions,
  buildMraiNativeFontAliasPackage,
  nextMraiNativeIndirectObjectNumber,
  planMraiNativeFontHistoryFromEvidence,
  planMraiNativeFontReuseHistory,
  planMraiNativePageFontExtensionHistory,
  planMraiNativePageFontExtensionMutation,
  PdfNativeTransactionError,
  prepareMraiNativeContentStreamTransaction,
  sha256PdfBytes,
  verifyMraiNativeContentStreamCandidate,
  type PdfNativeContentStreamTransactionRequest,
  type PdfNativeDuplicateSpan,
  type PdfNativeStructActualTextRevision,
  type PdfNativeFontExtensionHistoryObjectState,
  type PdfNativeFontAliasPackage,
  type PdfNativePageFontExtensionMutationPlan,
  type PdfNativePreparedContentStreamTransaction,
  type PdfNativePreparedFontExtensionHistory,
  type PdfNativeReusablePageFontExtension,
  type PdfNativeObjectPrecondition,
  ownedPdfSourceBytes,
} from "./engine/native/content-stream-transaction";
import {
  MRAI_NATIVE_FONT_EXTENSION_PACKAGE_SCHEMA_VERSION,
  PDF_NATIVE_FALLBACK_FONT_NAME,
  buildPdfNativeFontExtensionPackage,
  measurePdfNativeFontExtensionText,
  pdfNativeFontExtensionPackageDecoder,
  pdfNativeFontObjectWireBase64,
  pdfNativeFontObjectWireBytes,
  serializePdfNativeFontExtensionPackage,
  type PdfNativeFontExtensionPackage,
} from "./engine/native/font-extension-package";
import {
  createPdfNativeEmbeddedFontInstallation,
  PdfNativeEmbeddedFontError,
  type PdfNativeEmbeddedFontLibrary,
} from "./native-embedded-fonts";
import { readPdfPageFont, type PdfPageFont } from "./page-fonts";
import {
  loadEmbeddableFontManifest,
  verifyEmbeddableFontSlice,
  type EmbeddableFontManifest,
  type EmbeddableFontSlice,
} from "@mrnote/core/web/lib/office-fonts/embeddable-faces.ts";
import {
  planMraiNativeListTextParagraphFormat,
  pdfNativeParagraphPlanClearOfOtherText,
  planMraiNativeListTextPropertyReplacement,
  type PdfNativeListTextPropertyReplacementInput,
  type PdfNativeParagraphFormat,
  type PdfNativeTextFontOption,
  shiftedLineTail,
  MRAI_NATIVE_LIST_TEXT_SCHEMA_VERSION,
  MRAI_NATIVE_MULTI_RUN_TEXT_SCHEMA_VERSION,
  MRAI_NATIVE_PROJECTED_MULTI_RUN_LINE_PLAN_SCHEMA_VERSION,
  MRAI_NATIVE_LIST_TEXT_MISSING_GLYPH_MESSAGE,
  MRAI_NATIVE_DUPLICATE_SHOWS_MESSAGE,
  MRAI_NATIVE_GLYPHS_READ_AS_BLANKS_MESSAGE,
  planMraiNativeParagraphRewriteWithSubstitute,
  planMraiNativeStructActualTextForRequest,
  proveMraiNativeListTextSelectionForWriter,
  readMraiNativeEvidenceFaceStyle,
  wrapPdfNativeMeasuredGlyphsInFixedFrame,
  type PdfNativeListTextFidelityEvidence,
  type PdfNativeListTextSelectionEvidence,
  type PdfNativeListTextSelectionProof,
  type PdfNativeListTextWriterProof,
  type PdfNativeSingleLineFitHint,
} from "./engine/native/list-text-fidelity";
import {
  enclosingTextObject,
  mraiNativeShowOperatorsIn,
  mraiNativeTextAfterShowIsPlacedAfresh,
} from "./engine/native/text-object-delete";

export const PDF_NATIVE_FONT_EXTENSION_TEXT_FEATURE_ID =
  "PDF-EDT-TEXT-01" as const;
export const PDF_NATIVE_FONT_EXTENSION_TEXT_CONTROL_ID =
  "PDFCTL.RIBBON.EDIT.EDIT_CONTENT" as const;

const MAX_REPLACEMENT_CODE_UNITS = 2_000;

export interface PdfNativeFontExtensionSourceIdentity {
  readonly resourceId: string;
  readonly resourceVersionId: string;
}

export interface PdfNativeFontExtensionTextPlan {
  readonly sourceProof: PdfNativeListTextSelectionProof;
  readonly prepared: PdfNativePreparedContentStreamTransaction;
  readonly record: PdfSharedStateRecord;
}

export interface PdfNativeFontExtensionTextHistoryPlan {
  readonly prepared: PdfNativePreparedFontExtensionHistory;
  readonly record: PdfSharedStateRecord;
}

export interface PdfNativeFontReuseTextCandidatePlan {
  readonly sourceProof: PdfNativeListTextSelectionProof;
  readonly prepared: PdfNativePreparedContentStreamTransaction;
  readonly reuse: PdfNativeReusablePageFontExtension;
}

export interface PdfNativeFontReuseTextPlan
  extends PdfNativeFontReuseTextCandidatePlan {
  readonly record: PdfSharedStateRecord;
}

function invalid(message: string): never {
  throw new Error(message);
}

/** The objects the source proof already pinned -- the font's own indirect
 * /Widths, /Encoding, /W and descendant objects, and the page's other content
 * streams when it has several -- carried into a font transaction's
 * preconditions. Replay requires every one of them, and the proof is where
 * they were resolved; listing only the page's resource owner here left the
 * save to fail on exactly those fonts and pages. */
const LATIN1_TEXT = LATIN1_EXACT;

function withProvenResourcePreconditions(
  own: readonly PdfNativeObjectPrecondition[],
  proof: PdfNativeListTextWriterProof,
): readonly PdfNativeObjectPrecondition[] {
  const merged = new Map<string, PdfNativeObjectPrecondition>();
  for (const item of [
    ...own,
    ...(proof.request.untouchedObjects ?? []).filter(
      (item) => item.role === "resource" || item.role === "untouched-object",
    ),
  ]) {
    const key = `${item.objectNumber}:${item.generation}`;
    const existing = merged.get(key);
    if (existing && existing.expectedObjectSha256 !== item.expectedObjectSha256) {
      return invalid("Native font transaction resource preconditions disagree.");
    }
    if (!existing) merged.set(key, item);
  }
  return Object.freeze([...merged.values()]);
}














function fontSetDuplicateSpans(
  proof: PdfNativeListTextWriterProof,
  operatorSpan: Readonly<{ start: number; end: number }>,
  writesTheShowAlone: boolean,
): readonly PdfNativeDuplicateSpan[] | null {
  const copies = proof.request.duplicateSpans;
  if (!copies?.length) return null;
  const own = proof.evidence.textOperatorSpan;
  if (
    !writesTheShowAlone ||
    operatorSpan.start !== own.start ||
    operatorSpan.end !== own.end ||
    copies.some((copy) => copy.sourceStart !== undefined)
  ) {
    return invalid(MRAI_NATIVE_DUPLICATE_SHOWS_MESSAGE);
  }
  return Object.freeze(copies.map((copy) => Object.freeze({ start: copy.start, end: copy.end })));
}

/** The copies an origin wrote with its show (`duplicateSpans`), replayed
 * with it so its history re-derives the candidate the origin saved. */
function originDuplicateSpans(
  origin: Record<string, unknown>,
): { duplicateSpans?: readonly PdfNativeDuplicateSpan[] } {
  if (origin.duplicateSpans === undefined || origin.duplicateSpans === null) return {};
  if (!Array.isArray(origin.duplicateSpans) || origin.duplicateSpans.length === 0) {
    return invalid("Native font extension duplicate spans are invalid.");
  }
  return {
    duplicateSpans: Object.freeze(origin.duplicateSpans.map((entry) => {
      const span = record(entry);
      if (span.sourceStart !== undefined) {
        return invalid("Native font extension duplicate spans are invalid.");
      }
      return Object.freeze({
        start: nonnegativeInteger(span.start, "duplicate span start"),
        end: positiveInteger(span.end, "duplicate span end"),
      });
    })),
  };
}

function record(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return invalid("Native font extension history origin is invalid.");
  }
  return value as Record<string, unknown>;
}

function positiveInteger(value: unknown, label: string): number {
  if (!Number.isSafeInteger(value) || Number(value) < 1) {
    return invalid(`Native font extension ${label} is invalid.`);
  }
  return Number(value);
}

function nonnegativeInteger(value: unknown, label: string): number {
  if (!Number.isSafeInteger(value) || Number(value) < 0) {
    return invalid(`Native font extension ${label} is invalid.`);
  }
  return Number(value);
}

function sha(value: unknown, label: string): string {
  if (typeof value !== "string" || !/^[0-9a-f]{64}$/i.test(value)) {
    return invalid(`Native font extension ${label} is invalid.`);
  }
  return value.toLowerCase();
}

function bytes(value: unknown, label: string): Uint8Array {
  if (
    !Array.isArray(value) ||
    value.length < 1 ||
    value.length > 16 * 1024 * 1024 ||
    value.some(
      (item) =>
        !Number.isInteger(item) || Number(item) < 0 || Number(item) > 255,
    )
  ) {
    return invalid(`Native font extension ${label} is invalid.`);
  }
  return Uint8Array.from(value as number[]);
}

function base64Bytes(value: unknown): Uint8Array {
  if (
    typeof value !== "string" ||
    value.length < 4 ||
    value.length > 24 * 1024 * 1024 ||
    value.length % 4 !== 0 ||
    !/^[A-Za-z0-9+/]+={0,2}$/.test(value)
  ) {
    return invalid("Native font extension package bytes are invalid.");
  }
  let binary: string;
  try {
    binary = atob(value);
  } catch {
    return invalid("Native font extension package bytes are invalid.");
  }
  const decoded = Uint8Array.from(binary, (character) =>
    character.charCodeAt(0),
  );
  let canonical = "";
  for (let offset = 0; offset < decoded.length; offset += 0x8000) {
    canonical += String.fromCharCode(
      ...decoded.subarray(offset, offset + 0x8000),
    );
  }
  if (btoa(canonical) !== value) {
    return invalid("Native font extension package bytes are invalid.");
  }
  return decoded;
}

function encodeBase64Bytes(bytes: Uint8Array): string {
  let binary = "";
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000));
  }
  return btoa(binary);
}

function serializeFontMutationPackage(
  fontPackage: PdfNativeFontExtensionPackage | PdfNativeFontAliasPackage,
) {
  if (
    fontPackage.schemaVersion === MRAI_NATIVE_FONT_EXTENSION_PACKAGE_SCHEMA_VERSION
  ) {
    return serializePdfNativeFontExtensionPackage(fontPackage);
  }
  return Object.freeze({
    ...fontPackage,
    sourceFont: Object.freeze({ ...fontPackage.sourceFont }),
    sourceToUnicode: Object.freeze({ ...fontPackage.sourceToUnicode }),
    objects: Object.freeze(
      fontPackage.objects.map((object) =>
        Object.freeze({
          objectNumber: object.objectNumber,
          generation: object.generation,
          sha256: object.sha256,
          ...pdfNativeFontObjectWireBytes(encodeBase64Bytes(object.bytes)),
        }),
      ),
    ),
  });
}

/** Room beside a line, in its 1000-unit space, below which the line ends at
 * its frame's edge and grows to the left (a quarter of an em). */
const RIGHT_ANCHORED_ROOM_AT_1000 = 250;

function ascii(value: string): Uint8Array {
  return new TextEncoder().encode(value);
}

interface MixedFontSegment {
  /** Written by the source font (codes from the writer proof's table) or,
   * when false, by the extension font. */
  readonly source: boolean;
  readonly text: string;
  readonly codes: readonly number[];
  readonly advance: number;
  /** A source segment's pieces: the run's own units it keeps, or a
   * character written from the table, each with its codes and page advance. */
  readonly units?: readonly (readonly [string, readonly number[], number])[];
}

/** Split a replacement into the characters the source font writes itself
 * and the ones only the extension font has, one segment per stretch. Null
 * when the run is all one or all the other, or the proof carries no table. */
/** Whether the source font writes every character of `text` back exactly. */
function allSourceEncodable(
  hint: PdfNativeSingleLineFitHint | undefined,
  text: string,
): boolean {
  if (text === "") return true;
  if (!hint?.encodings) return false;
  const encodings = new Set(
    hint.encodings.filter(([character]) => [...character].length === 1).map(([character]) => character),
  );
  const advances = new Set(hint.advances.map(([character]) => character));
  return [...text].every((character) => encodings.has(character) && advances.has(character));
}

/** Whether the run's font is the face this route embeds: the bundled
 * fallback, which every writer here names MRNoteNotoSansSC (the name the
 * server validates), subset or whole. */
async function runIsInExtensionFace(
  sourceBytes: Uint8Array,
  fidelity: PdfNativeListTextFidelityEvidence,
): Promise<boolean> {
  if (fidelity.resources.fontInline) return false;
  const font = await inspectMraiNativeObject(sourceBytes, fidelity.resources.font);
  const names = [...LATIN1_EXACT.decode(font.objectBytes).matchAll(/\/BaseFont\s*\/([^\s/\[\]<>(){}%]+)/g)];
  return names.length === 1 && names[0][1].replace(/^[A-Z]{6}\+/, "") === PDF_NATIVE_FALLBACK_FONT_NAME;
}

type SourceUnit = readonly [string, readonly number[], number];

/** Where the run's own units go in the replacement: `at` a replacement
 * index, a group of units standing for the replacement characters from
 * there (`consume` of them); `before` an index, units the reader's text
 * never showed (white space it trimmed from the line's ends). */
interface KeptSourceUnits {
  readonly at: ReadonlyMap<number, Readonly<{ units: readonly SourceUnit[]; consume: number }>>;
  readonly before: ReadonlyMap<number, readonly SourceUnit[]>;
}

const READER_WHITE_SPACE = /^\s+$/u;

/**
 * The run's own units the replacement keeps. The replacement is the text
 * the reader shows, edited: the reader trims white space from a line's ends
 * and reads a run of white space as one space, so the run is matched as the
 * reader reads it -- each white-space run one space, the ends' white space
 * nothing -- and the characters the edit leaves alone keep every code they
 * were drawn with, the white space the reader folded or trimmed included
 * (a glyph its ToUnicode misnames a space -- the owner's lab script named
 * its capital A, its colon and its quotes so -- is one of those). The head
 * and tail match one to one, the stretch between by the longest common
 * subsequence; a group is kept only when every character of it stays, in
 * one piece. Null without units or past the size the match is made for.
 */
function keptSourceUnits(
  units: PdfNativeSingleLineFitHint["sourceUnits"],
  replacementText: string,
): KeptSourceUnits | null {
  if (!units?.length) return null;
  // The run as the reader reads it: groups of units, each with its text.
  const groups: Array<{ units: SourceUnit[]; text: string; space: boolean }> = [];
  for (const unit of units) {
    const space = READER_WHITE_SPACE.test(unit[0]);
    const last = groups.at(-1);
    if (space && last?.space) {
      last.units.push(unit);
      continue;
    }
    groups.push({ units: [unit], text: space ? " " : unit[0], space });
  }
  const firstContent = groups.findIndex((group) => !group.space);
  if (firstContent < 0) return null;
  let lastContent = groups.length - 1;
  while (groups[lastContent]!.space) lastContent -= 1;
  const leading = groups.slice(0, firstContent).flatMap((group) => group.units);
  const trailing = groups.slice(lastContent + 1).flatMap((group) => group.units);
  const shown = groups.slice(firstContent, lastContent + 1);
  const originalChars: string[] = [];
  const firstCharOfGroup: number[] = [];
  for (const group of shown) {
    firstCharOfGroup.push(originalChars.length);
    for (const character of group.text) originalChars.push(character);
  }
  const replacementChars = [...replacementText];
  if (originalChars.length > 2_000 || replacementChars.length > 2_000) return null;
  const matchOfOriginal = new Int32Array(originalChars.length).fill(-1);
  let head = 0;
  while (
    head < originalChars.length && head < replacementChars.length &&
    originalChars[head] === replacementChars[head]
  ) {
    matchOfOriginal[head] = head;
    head += 1;
  }
  let tail = 0;
  while (
    head + tail < originalChars.length && head + tail < replacementChars.length &&
    originalChars[originalChars.length - 1 - tail] === replacementChars[replacementChars.length - 1 - tail]
  ) {
    matchOfOriginal[originalChars.length - 1 - tail] = replacementChars.length - 1 - tail;
    tail += 1;
  }
  const originalEnd = originalChars.length - tail;
  const replacementEnd = replacementChars.length - tail;
  const columns = replacementEnd - head + 1;
  const lcs = new Uint16Array((originalEnd - head + 1) * columns);
  for (let i = originalEnd - 1; i >= head; i -= 1) {
    for (let j = replacementEnd - 1; j >= head; j -= 1) {
      const cell = (i - head) * columns + (j - head);
      lcs[cell] = originalChars[i] === replacementChars[j]
        ? lcs[cell + columns + 1] + 1
        : Math.max(lcs[cell + columns], lcs[cell + 1]);
    }
  }
  for (let i = head, j = head; i < originalEnd && j < replacementEnd; ) {
    const cell = (i - head) * columns + (j - head);
    if (originalChars[i] === replacementChars[j]) {
      matchOfOriginal[i] = j;
      i += 1;
      j += 1;
    } else if (lcs[cell + columns] >= lcs[cell + 1]) {
      i += 1;
    } else {
      j += 1;
    }
  }
  const at = new Map<number, Readonly<{ units: readonly SourceUnit[]; consume: number }>>();
  const keptGroup = shown.map((group, groupIndex) => {
    const start = firstCharOfGroup[groupIndex]!;
    const count = [...group.text].length;
    for (let offset = 0; offset < count; offset += 1) {
      const match = matchOfOriginal[start + offset]!;
      if (match < 0 || (offset > 0 && match !== matchOfOriginal[start + offset - 1]! + 1)) return false;
    }
    at.set(matchOfOriginal[start]!, Object.freeze({ units: Object.freeze(group.units), consume: count }));
    return true;
  });
  const before = new Map<number, readonly SourceUnit[]>();
  // The ends' white space goes back beside the character the reader showed
  // next to it, while the edit leaves that character where it was.
  if (leading.length && keptGroup[0] && matchOfOriginal[0] === 0) {
    before.set(0, Object.freeze(leading));
  }
  const lastShown = shown.length - 1;
  if (trailing.length && keptGroup[lastShown]) {
    const end = matchOfOriginal[firstCharOfGroup[lastShown]!]! + [...shown[lastShown]!.text].length;
    before.set(end, Object.freeze([...(before.get(end) ?? []), ...trailing]));
  }
  return Object.freeze({ at, before });
}

function splitReplacementBySourceFont(
  hint: PdfNativeSingleLineFitHint | undefined,
  replacementText: string,
): readonly MixedFontSegment[] | null {
  if (!hint?.encodings) return null;
  const encodings = new Map(
    hint.encodings.filter(([character]) => [...character].length === 1),
  );
  const advances = new Map(hint.advances);
  const segments: Array<{
    source: boolean;
    text: string;
    codes: number[];
    advance: number;
    units: SourceUnit[];
  }> = [];
  const push = (source: boolean, text: string, codes: readonly number[], advance: number) => {
    const last = segments.at(-1);
    if (last && last.source === source) {
      last.text += text;
      if (source) {
        last.codes.push(...codes);
        last.advance += advance;
        last.units.push(Object.freeze([text, Object.freeze([...codes]), advance] as const));
      }
      return;
    }
    segments.push({
      source,
      text,
      codes: source ? [...codes] : [],
      advance: source ? advance : 0,
      units: source ? [Object.freeze([text, Object.freeze([...codes]), advance] as const)] : [],
    });
  };
  // The text the edit leaves alone keeps the run's own codes: re-encoded
  // from the font's names, a glyph its ToUnicode misnames came back as
  
  // space, and every space of the rewritten line was drawn an A).
  const kept = keptSourceUnits(hint.sourceUnits, replacementText);
  const replacementChars = [...replacementText];
  const emitBefore = (index: number) => {
    for (const unit of kept?.before.get(index) ?? []) push(true, unit[0], unit[1], unit[2]);
  };
  for (let index = 0; index < replacementChars.length; ) {
    emitBefore(index);
    const group = kept?.at.get(index);
    if (group) {
      for (const unit of group.units) push(true, unit[0], unit[1], unit[2]);
      index += group.consume;
      continue;
    }
    const character = replacementChars[index]!;
    const codes = encodings.get(character);
    const advance = advances.get(character);
    const source = codes !== undefined && advance !== undefined;
    push(source, character, codes ?? [], advance ?? 0);
    index += 1;
  }
  emitBefore(replacementChars.length);
  if (!segments.some((segment) => segment.source) || !segments.some((segment) => !segment.source)) {
    return null;
  }
  return Object.freeze(
    segments.map((segment) =>
      Object.freeze({
        ...segment,
        codes: Object.freeze(segment.codes),
        units: Object.freeze(segment.units),
      }),
    ),
  );
}

interface MixedShowLayout {
  /** Advance of the whole show: source segments by their own font's
   * metrics, extension segments by the installed glyphs. */
  readonly advance: number;
  /** Glyph range of each extension segment, null for a source segment. */
  readonly ranges: readonly (readonly [number, number] | null)[];
  readonly glyphCount: number;
  readonly widthAt1000: number;
}

/** Lay the extension glyphs over the segments; null when a glyph spans a
 * segment boundary (a ligature across the split) or glyphs are left over. */
function mixedShowLayout(
  mixed: readonly MixedFontSegment[],
  glyphs: Awaited<ReturnType<typeof measurePdfNativeFontExtensionText>>["glyphs"],
  fontSize: number,
): MixedShowLayout | null {
  const ranges: Array<readonly [number, number] | null> = [];
  let glyphCursor = 0;
  let advance = 0;
  let widthAt1000 = 0;
  for (const segment of mixed) {
    if (segment.source) {
      advance += segment.advance;
      ranges.push(null);
      continue;
    }
    const start = glyphCursor;
    let text = "";
    while (text.length < segment.text.length && glyphCursor < glyphs.length) {
      const glyph = glyphs[glyphCursor];
      text += glyph.text;
      advance += (glyph.advanceAt1000 * fontSize) / 1_000;
      widthAt1000 += glyph.advanceAt1000;
      glyphCursor += 1;
    }
    if (text !== segment.text) return null;
    ranges.push(Object.freeze([start, glyphCursor] as const));
  }
  if (glyphCursor !== glyphs.length) return null;
  return Object.freeze({
    advance,
    ranges: Object.freeze(ranges),
    glyphCount: glyphs.length,
    widthAt1000,
  });
}

/** One unit a framed line wraps by: a character the source font writes, or
 * one glyph of the extension font. Each carries the width it draws in the
 * line's own 1000-unit space, so a line may hold both fonts and still be
 * measured against the frame the way a single-font line is. */
interface FrameUnit {
  readonly text: string;
  readonly advanceAt1000: number;
  readonly source: boolean;
  /** Source codes for a source unit, extension glyph hex for the other. */
  readonly hex: string;
}

/** Lay the replacement out as frame units. Null when a glyph of the
 * extension font spans a segment boundary (a ligature across the split) or
 * glyphs are left over, which leaves the two fonts unalignable. */
function frameUnits(
  mixed: readonly MixedFontSegment[],
  glyphs: Awaited<ReturnType<typeof measurePdfNativeFontExtensionText>>["glyphs"],
  encodedHexOfGlyph: (index: number) => string,
  sourceCharacter: (
    character: string,
  ) => Readonly<{ codes: readonly number[]; advanceAt1000: number }> | null,
  characterSpacingAt1000: number,
  /** A source unit's page advance in the line's 1000-unit space. */
  sourceAdvanceAt1000: (advance: number) => number,
): readonly FrameUnit[] | null {
  const units: FrameUnit[] = [];
  let glyphCursor = 0;
  for (const segment of mixed) {
    if (segment.source && segment.units?.length) {
      // The pieces the split kept, each with its own codes: the run's own
      // where the edit left the text alone.
      for (const [text, codes, advance] of segment.units) {
        units.push(
          Object.freeze({
            text,
            advanceAt1000: sourceAdvanceAt1000(advance),
            source: true,
            hex: hexOf([...codes]),
          }),
        );
      }
      continue;
    }
    if (segment.source) {
      for (const character of segment.text) {
        const written = sourceCharacter(character);
        if (written === null) return null;
        units.push(
          Object.freeze({
            text: character,
            advanceAt1000: written.advanceAt1000,
            source: true,
            hex: hexOf([...written.codes]),
          }),
        );
      }
      continue;
    }
    let text = "";
    while (text.length < segment.text.length && glyphCursor < glyphs.length) {
      const glyph = glyphs[glyphCursor];
      text += glyph.text;
      units.push(
        Object.freeze({
          text: glyph.text,
          advanceAt1000: glyph.advanceAt1000 + characterSpacingAt1000,
          source: false,
          hex: encodedHexOfGlyph(glyphCursor),
        }),
      );
      glyphCursor += 1;
    }
    if (text !== segment.text) return null;
  }
  return glyphCursor === glyphs.length ? Object.freeze(units) : null;
}

function hexOf(codes: readonly number[]): string {
  return codes.map((code) => code.toString(16).toUpperCase().padStart(2, "0")).join("");
}

function concat(parts: readonly Uint8Array[]): Uint8Array {
  const output = new Uint8Array(
    parts.reduce((total, part) => total + part.length, 0),
  );
  let offset = 0;
  for (const part of parts) {
    output.set(part, offset);
    offset += part.length;
  }
  return output;
}

function formatNumber(value: number): string {
  if (!Number.isFinite(value) || value <= 0 || value > 1_000) {
    return invalid("Invalid native font extension size.");
  }
  return value
    .toFixed(6)
    .replace(/\.0+$/u, "")
    .replace(/(\.\d*?)0+$/u, "$1");
}

/** A TJ offset (thousandths of the size) that moves the pen back before a
 * line's first glyph. */
function formatLeadOffset(value: number): string {
  if (!Number.isFinite(value) || value <= 0 || value > 1_000_000) {
    return invalid("Invalid native font extension line offset.");
  }
  return value.toFixed(3).replace(/\.?0+$/u, "");
}

async function resourceName(input: {
  sourceSha256: string;
  pageNumber: number;
  targetObjectNumber: number;
  operatorStart: number;
  replacementText: string;
}): Promise<string> {
  const material = new TextEncoder().encode(
    [
      "mrai.native-font-extension.v1",
      input.sourceSha256,
      input.pageNumber,
      input.targetObjectNumber,
      input.operatorStart,
      input.replacementText,
    ].join("\0"),
  );
  const digest = new Uint8Array(
    await crypto.subtle.digest("SHA-256", material),
  );
  return `MRAI_NF_${[...digest]
    .slice(0, 12)
    .map((value) => value.toString(16).padStart(2, "0"))
    .join("")}`;
}






















function positionedRunGroupRest(
  streamBytes: Uint8Array,
  fidelity: PdfNativeListTextFidelityEvidence,
): Uint8Array {
  const runs = fidelity.positionedRunGroup?.runs.map((run) => run.evidence) ?? [];
  const first = runs[0];
  const last = runs.at(-1);
  const groupEnd = fidelity.lineShift?.showEnd ?? fidelity.textOperatorSpan.end;
  if (
    runs.length < 2 ||
    !first ||
    !last ||
    first.textOperatorSpan.start !== fidelity.textOperatorSpan.start ||
    last.textOperatorSpan.end !== groupEnd ||
    runs.some(
      (run) =>
        run.readerChunk !== undefined ||
        run.emptiedShow !== undefined ||
        run.positionedRunGroup !== undefined ||
        run.consolidatedRuns !== undefined ||
        run.resourceName !== first.resourceName ||
        Math.abs(run.fontSize - first.fontSize) > 0.000001,
    ) ||
    (/^\s/u.test(first.originalText) && !/^\s/u.test(fidelity.originalDisplayText))
  ) {
    return invalid("Native font extension cannot set this run group in its first run.");
  }
  const source = LATIN1_TEXT.decode(streamBytes);
  const shows = mraiNativeShowOperatorsIn(source, first.textOperatorSpan.start, groupEnd);
  if (
    !shows ||
    shows.length !== runs.length ||
    shows.some(
      (show, index) =>
        show.start !== runs[index].textOperatorSpan.start ||
        show.end !== runs[index].textOperatorSpan.end,
    ) ||
    (shows[0].emptyForm !== "<> Tj" && shows[0].emptyForm !== "[] TJ")
  ) {
    return invalid("Native font extension run group shows changed.");
  }
  const parts: Uint8Array[] = [];
  let cursor = first.textOperatorSpan.end;
  for (const [index, show] of shows.entries()) {
    let objectEnd: number;
    try {
      objectEnd = enclosingTextObject(source, show.start).end;
    } catch {
      return invalid("Native font extension run group shows changed.");
    }
    // The first run now draws the whole line and every later run nothing:
    // text its object draws on from the end of either must be placed
    // afresh, or be another run of the group.
    if (!mraiNativeTextAfterShowIsPlacedAfresh(source, show.end, groupEnd, objectEnd)) {
      return invalid("Native font extension run group would move the text after it.");
    }
    if (index === 0) continue;
    parts.push(streamBytes.slice(cursor, show.start), ascii(show.emptyForm));
    cursor = show.end;
  }
  return concat(parts);
}

/** The advance a single run's show takes once written, per plan. */
const singleRunReplacementAdvances = new WeakMap<PdfNativeFontExtensionTextPlan, number>();















function positionedRunGroupEditRun(
  selection: PdfNativeListTextSelectionEvidence,
  fidelity: PdfNativeListTextFidelityEvidence,
  replacementText: string,
): Readonly<{
  selection: PdfNativeListTextSelectionEvidence;
  replacementText: string;
  followedByInk: boolean;
}> | null {
  const runs = fidelity.positionedRunGroup?.runs ?? [];
  const before = fidelity.originalDisplayText;
  if (runs.length < 2 || fidelity.readerChunk !== undefined || before === replacementText) return null;
  const displays = runs.map((run) => run.evidence.originalDisplayText);
  const starts: number[] = [];
  const lengths: number[] = [];
  let offset = 0;
  for (const display of displays) {
    let start = before.indexOf(display, offset);
    let length = display.length;
    // A closing run of nothing but white space the reader leaves off the
    // item (Word ends a title with a space in an object of its own).
    if (start < 0 && !display.trim() && !before.slice(offset).trim()) {
      start = offset;
      length = 0;
    }
    if (start < 0 || before.slice(offset, start).trim()) return null;
    starts.push(start);
    lengths.push(length);
    offset = start + length;
  }
  if (before.slice(offset).trim()) return null;
  let head = 0;
  while (head < before.length && head < replacementText.length && before[head] === replacementText[head]) head += 1;
  let tail = 0;
  while (
    tail < before.length - head &&
    tail < replacementText.length - head &&
    before[before.length - 1 - tail] === replacementText[replacementText.length - 1 - tail]
  ) {
    tail += 1;
  }
  const changedEnd = before.length - tail;
  const runIndex = starts.findIndex(
    (start, index) => lengths[index] > 0 && start <= head && changedEnd <= start + lengths[index],
  );
  if (runIndex < 0) return null;
  const run = runs[runIndex].evidence;
  const item = run.selection?.item;
  const text = replacementText.slice(
    starts[runIndex],
    starts[runIndex] + lengths[runIndex] + replacementText.length - before.length,
  );
  if (
    !item ||
    item.str !== displays[runIndex] ||
    run.readerChunk !== undefined ||
    run.emptiedShow !== undefined ||
    !text.trim() ||
    text === displays[runIndex]
  ) {
    return null;
  }
  return Object.freeze({
    selection: Object.freeze({
      sourceSha256: selection.sourceSha256,
      selectionRevision: selection.selectionRevision,
      pageNumber: selection.pageNumber,
      startOffset: 0,
      endOffset: item.str.length,
      item,
    }),
    replacementText: text,
    followedByInk: displays.slice(runIndex + 1).some((display) => display.trim() !== ""),
  });
}

function derivedObjects(
  prepared: PdfNativePreparedContentStreamTransaction,
): readonly Readonly<{
  objectNumber: number;
  generation: number;
  offset: number;
  sha256: string;
}>[] {
  const mutation = prepared.pageFontExtensionMutation;
  if (!mutation) return invalid("Native font extension receipt is missing.");
  // A font installed in an isolated Form rides on the clone chain: every
  // object that transaction wrote is the revision set.
  const isolation = prepared.formInstanceIsolation;
  if (isolation) {
    return Object.freeze(
      [...isolation.objects, ...pdfNativePdfaIdentificationRevisedObjects(prepared)].sort(
        (left, right) =>
          left.objectNumber - right.objectNumber ||
          left.generation - right.generation,
      ),
    );
  }
  // An unshared form's receipt names the edited form as its "page": one
  // object, listed once.
  const seen = new Set<string>();
  return Object.freeze(
    [
      {
        objectNumber: prepared.request.identity.objectNumber,
        generation: prepared.request.identity.generation,
        offset: prepared.changedObjectOffset,
        sha256: prepared.changedObjectSha256,
      },
      ...mutation.fontObjects,
      mutation.resources,
      mutation.page,
      
      ...(prepared.structActualText ?? []),
      // The catalog's packet, without its PDF/A identification.
      ...pdfNativePdfaIdentificationRevisedObjects(prepared),
    ].filter((item) => {
      const key = `${item.objectNumber}:${item.generation}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    }).sort(
      (left, right) =>
        left.objectNumber - right.objectNumber ||
        left.generation - right.generation,
    ),
  );
}






async function withStructActualText(
  sourceBytes: Uint8Array,
  pageNumber: number,
  request: PdfNativeContentStreamTransactionRequest,
  packages: readonly Pick<PdfNativeFontExtensionPackage, "resourceName" | "objects">[],
  /** The edited line as the reader reads it, before and after: an element
   * whose content's fonts cannot be read is written by it. */
  lineTexts?: Readonly<{ before: string; after: string }>,
): Promise<PdfNativeContentStreamTransactionRequest> {
  const installedFonts = new Map<string, (codes: Uint8Array) => string | null>();
  for (const fontPackage of packages) {
    const decode = pdfNativeFontExtensionPackageDecoder(fontPackage);
    if (decode) installedFonts.set(fontPackage.resourceName, decode);
  }
  return planMraiNativeStructActualTextForRequest({
    sourceBytes, request, pageNumber, installedFonts, ...(lineTexts ? { lineTexts } : {}),
  });
}

/** The run route: one exact source run or fixed frame, the characters its
 * own font writes kept in it. */
async function planPdfNativeFontExtensionRunTextEdit(input: {
  readonly sourceBytes: Uint8Array;
  readonly source: PdfNativeFontExtensionSourceIdentity;
  readonly selection: PdfNativeListTextSelectionEvidence;
  readonly replacementText: string;
  readonly fontBytes: Uint8Array | ArrayBuffer;
  readonly transactionId: string;
  readonly baseSessionRevision: number;
  /** Second attempt: prove the run with the line's later runs moving, when
   * the free space beside it could not hold the extension glyphs. */
  readonly preferLineShift?: boolean;
  /** The run of a positioned run group the edit falls inside was tried (or
   * this is that run): the group is set in its first run. */
  readonly groupRunTried?: boolean;
  /** Formats asked of the paragraph with the text: only the paragraph route
   * sets them (planPdfNativeFontExtensionTextEdit). */
  readonly paragraphFormat?: Readonly<{ format: PdfNativeParagraphFormat;
    fontOptions?: readonly PdfNativeTextFontOption[] }>;
  /** The frame the paragraph is set in (a frame the reader resized, the
   * text box the words were typed into): only the paragraph route sets
   * words in a frame. */
  readonly paragraphFrame?: Readonly<{ right?: number; lineStep?: number; sizeScale?: number }>;
  /** Where the editor's rows start in the new text (native-text-row-starts):
   * the paragraph route keeps words typed at the end of a row on it. */
  readonly rowStarts?: readonly number[];
}): Promise<PdfNativeFontExtensionTextPlan> {
  if (
    !input.source.resourceId ||
    !input.source.resourceVersionId ||
    !input.transactionId ||
    !Number.isSafeInteger(input.baseSessionRevision) ||
    input.baseSessionRevision < 0 ||
    !input.replacementText ||
    input.replacementText.length > MAX_REPLACEMENT_CODE_UNITS ||
    /[\u0000-\u001f\u007f]/u.test(input.replacementText)
  ) {
    return invalid("Invalid native font extension request.");
  }
  const sourceBytes = ownedPdfSourceBytes(input.sourceBytes);
  const sourceSha256 = await sha256PdfBytes(sourceBytes);
  if (sourceSha256 !== input.selection.sourceSha256.toLowerCase()) {
    return invalid("Native font extension source is stale.");
  }
  const reusable = await inspectMraiNativeReusablePageFontExtension({
    sourceBytes,
    pageNumber: input.selection.pageNumber,
    replacementText: input.replacementText,
    fontBytes: input.fontBytes,
  });
  if (
    reusable &&
    reusable.currentToUnicode.objectSha256 ===
      reusable.replacementToUnicode.objectSha256
  ) {
    try {
      return await planPdfNativeFontReuseTextEdit({ ...input, reusable });
    } catch {
      // The direct reuse rewrites one plain show. A run it cannot take -- a
      // fake-bold heading drawn twice glyph by glyph, a fixed frame, a reader
      // chunk -- is still set below, through an alias of the installed font
      // rather than a second copy of its program.
    }
  }
  const sourceProof = await proveMraiNativeListTextSelectionForWriter({
    sourceBytes,
    selection: input.selection,
    ...(input.preferLineShift ? { preferLineShift: true } : {}),
  });
  // A line of several text objects whose edit falls inside one of them: that
  // run takes it alone, under its own state, and the others keep their bytes
  // (positionedRunGroupEditRun). It is the write when nothing else on the line
  // has to move for it -- no inked run follows, the run's own proof moves the
  // runs after it, or its width stays -- and otherwise the one the group's
  // first run cannot take whole; a gap the narrower run leaves is a notice,
  // never a refusal (docs/pdf/edit-gates.md).
  const groupRun =
    sourceProof.evidence.positionedRunGroup && !input.groupRunTried
      ? positionedRunGroupEditRun(input.selection, sourceProof.evidence, input.replacementText)
      : null;
  if (groupRun) {
    const runPlan = await planPdfNativeFontExtensionRunTextEdit({
      ...input,
      selection: groupRun.selection,
      replacementText: groupRun.replacementText,
      preferLineShift: false,
      groupRunTried: true,
    }).catch(() => null);
    const runEvidence = runPlan?.sourceProof.evidence;
    const runAdvance = runPlan ? singleRunReplacementAdvances.get(runPlan) : undefined;
    if (
      runPlan &&
      runEvidence &&
      (!groupRun.followedByInk ||
        runEvidence.lineShift !== undefined ||
        (runAdvance !== undefined &&
          Math.abs(runAdvance - runEvidence.originalAdvance) <= runEvidence.fontSize * 0.01))
    ) {
      return runPlan;
    }
    try {
      return await planPdfNativeFontExtensionRunTextEdit({ ...input, groupRunTried: true });
    } catch (error) {
      if (runPlan) return runPlan;
      throw error;
    }
  }
  const fidelity = sourceProof.evidence;
  const linePlan = fidelity.linePlan;
  const visualLineIndexes =
    linePlan?.lines.map((line, index) => line.lineIndex ?? index) ?? [];
  const visualLinesAreOrdered =
    visualLineIndexes.length >= 2 &&
    visualLineIndexes[0] === 0 &&
    visualLineIndexes.every(
      (lineIndex, index) =>
        lineIndex === (visualLineIndexes[index - 1] ?? 0) ||
        lineIndex === (visualLineIndexes[index - 1] ?? -1) + 1,
    );
  // One renderer item projected onto several shows is a frame of one visual
  // line: the whole replacement goes into its first show and the rest are
  // emptied, the way a fixed frame refills its lines.
  const fixedFrame =
    fidelity.wrapping === "fixed-frame-multiline" &&
    linePlan !== null &&
    linePlan.lines.length >= 2 &&
    (fidelity.selection?.items !== undefined ||
      linePlan.schemaVersion ===
        MRAI_NATIVE_PROJECTED_MULTI_RUN_LINE_PLAN_SCHEMA_VERSION) &&
    visualLinesAreOrdered &&
    linePlan.lines.every(
      (line, index) =>
        line.layoutState !== undefined &&
        (line.operatorKind === "Tj" ||
          line.operatorKind === "Tj-literal" ||
          line.operatorKind === "'" ||
          line.operatorKind === "'-literal") &&
        (index === 0 ||
          visualLineIndexes[index] !== visualLineIndexes[index - 1] ||
          line.operatorKind === "Tj" ||
          line.operatorKind === "Tj-literal"),
    );
  const singleRun =
    fidelity.schemaVersion === MRAI_NATIVE_LIST_TEXT_SCHEMA_VERSION &&
    fidelity.wrapping === "same-line-only" &&
    fidelity.linePlan === null &&
    fidelity.selection?.items === undefined;
  // How far this one run may grow: its own width, or the free space beside
  // it the source proof measured with replay's rule.
  const singleRunCapacity = Math.max(
    fidelity.originalAdvance,
    sourceProof.singleLineFit?.capacity ?? 0,
  );
  // The selection may be one reader item of a longer show (the reader cuts
  // a show at wide kerning gaps): the item is set anew and the codes around
  // it keep the source font, as one mixed show over the whole run.
  const chunk = fidelity.readerChunk
    ? sourceProof.singleLineFit?.readerChunkPrefix !== undefined &&
      sourceProof.singleLineFit.readerChunkSuffix !== undefined
      ? {
          prefix: sourceProof.singleLineFit.readerChunkPrefix,
          suffix: sourceProof.singleLineFit.readerChunkSuffix,
        }
      : null
    : { prefix: "", suffix: "" };
  if (chunk === null) {
    return invalid("Native font extension reader chunk context is missing.");
  }
  const wholeReplacementText = chunk.prefix + input.replacementText + chunk.suffix;
  if (
    (fidelity.schemaVersion !== MRAI_NATIVE_LIST_TEXT_SCHEMA_VERSION &&
      fidelity.schemaVersion !== MRAI_NATIVE_MULTI_RUN_TEXT_SCHEMA_VERSION) ||
    (!singleRun && !fixedFrame) ||
    // A line consolidated into one show, or several positioned lines, is
    // no run the first show can take (positionedRunGroupRest).
    fidelity.consolidatedRuns !== undefined ||
    fidelity.positionedLines !== undefined ||
    fidelity.propertyMutation !== undefined ||
    fidelity.paragraphReflow !== undefined ||
    (fidelity.readerChunk !== undefined && !singleRun) ||
    fidelity.originalText !== fidelity.replacementText ||
    (singleRun &&
      fidelity.originalCodeBytes.join(",") !==
        fidelity.replacementCodeBytes.join(",")) ||
    fidelity.originalText === wholeReplacementText ||
    wholeReplacementText.length > MAX_REPLACEMENT_CODE_UNITS
  ) {
    return invalid(
      "Native font extension requires one exact source run or fixed frame.",
    );
  }
  // The codes around a reader chunk stay exactly as written: every one of
  // them must be a character the source font writes back.
  if (
    fidelity.readerChunk &&
    !(
      allSourceEncodable(sourceProof.singleLineFit, chunk.prefix) &&
      allSourceEncodable(sourceProof.singleLineFit, chunk.suffix)
    )
  ) {
    return invalid("Native font extension cannot keep the codes around the reader chunk.");
  }
  const page = await inspectMraiNativePageContentGraph(
    sourceBytes,
    input.selection.pageNumber,
  );
  // The proof already names the one stream the edit lives in; a page drawn
  // from several streams keeps the others pinned as untouched objects. A
  // page whose streams straddle a text object was planned as one
  // consolidated stream; the same virtual stream is inspected here.
  const consolidation = sourceProof.request.pageContentConsolidation ?? null;
  // Text drawn by a Form XObject was planned on the form itself, or on an
  // isolated clone when the form is shared; the font then goes into that
  // form's own resources rather than the page's.
  const isolation = sourceProof.request.formInstanceIsolation ?? null;
  const isolatedLeaf = isolation?.chain.at(-1) ?? null;
  const targetIsPageStream = page.contentStreams.some(
    (candidate) =>
      candidate.objectNumber === sourceProof.target.objectNumber &&
      candidate.generation === sourceProof.target.generation,
  );
  const formTarget =
    isolatedLeaf ??
    (!isolation && !consolidation && !targetIsPageStream ? sourceProof.target : null);
  if (
    isolation
      ? consolidation !== null ||
        !isolatedLeaf ||
        isolatedLeaf.objectNumber !== sourceProof.target.objectNumber ||
        isolatedLeaf.generation !== sourceProof.target.generation ||
        isolation.page.objectNumber !== page.page.objectNumber ||
        isolation.page.generation !== page.page.generation
      : consolidation
      ? page.contentStreams.length !== consolidation.sourceStreams.length ||
        page.contentStreams.some(
          (candidate, index) =>
            candidate.objectNumber !==
              consolidation.sourceStreams[index].objectNumber ||
            candidate.generation !== consolidation.sourceStreams[index].generation,
        )
      : false
  ) {
    return invalid("Native font extension stream ownership is ambiguous.");
  }
  const stream = consolidation
    ? await inspectMraiNativeConsolidatedContentStream(sourceBytes, consolidation)
    : await inspectMraiNativeContentStream(sourceBytes, sourceProof.target);
  // A line of several text objects: the replacement goes into its first run
  // and every later run is emptied in its own object.
  const runGroupRest =
    singleRun && fidelity.positionedRunGroup
      ? positionedRunGroupRest(stream.streamBytes, fidelity)
      : null;
  const name = await resourceName({
    sourceSha256,
    pageNumber: page.pageNumber,
    targetObjectNumber: sourceProof.target.objectNumber,
    operatorStart: fidelity.textOperatorSpan.start,
    replacementText: input.replacementText,
  });
  // The consolidated stream takes the first free number; the font follows.
  const nextObjectNumber = nextMraiNativeIndirectObjectNumber(sourceBytes);
  if (consolidation && consolidation.stream.objectNumber !== nextObjectNumber) {
    return invalid("Consolidated content stream allocation changed.");
  }
  // The clones of an isolated Form chain take the first free numbers too.
  const firstObjectNumber =
    nextObjectNumber + (consolidation ? 1 : isolation ? isolation.chain.length : 0);
  // A run keeps every character its own font can write and sets only the
  // others in the extension font, so it grows by those few glyphs rather
  // than by a whole line in a wider face. A framed line is laid out the
  // same way: a label that gains one substituted letter still fits its
  // cell, where the whole label reset in a wider face would not.
  let mixed = !reusable
    ? splitReplacementBySourceFont(sourceProof.singleLineFit, wholeReplacementText)
    : null;
  // A run already in the extension's own face -- a new text box's first
  // words, written in a subset holding only the letters typed so far -- is
  // set whole in it: kept split, the box's first letter stayed a run of its
  // own beside the rest, two pieces in Edit Content that dragged apart
  
  if (mixed && !fidelity.readerChunk && (await runIsInExtensionFace(sourceBytes, fidelity))) {
    mixed = null;
  }
  // A reader chunk is only ever set as a mixed show: the codes around the
  // item are not this font's to change.
  if (fidelity.readerChunk && !mixed) {
    return invalid("Native font extension cannot set a reader chunk as a whole run.");
  }
  // A glyph the run's font only calls a space reads as a blank where the
  // page shows a letter (the hint's inkedBlankCodes). The run keeps each one
  // it drew -- the units the edit leaves alone are laid out as they are --
  // and where it would not (the whole run set in the extension face, or the
  // edit taking the blank the glyph was read as), nothing is written.
  const inkedBlanks = sourceProof.singleLineFit?.inkedBlankCodes;
  const keepsGlyphsReadAsBlanks = (segments: readonly MixedFontSegment[] | null) => {
    if (!inkedBlanks?.length) return;
    const inked = new Set(inkedBlanks);
    const kept = (segments ?? []).reduce((total, segment) => total + (segment.units ?? [])
      .filter((unit) => inked.has(unit[1].reduce((code, byte) => (code << 8) | byte, 0))).length, 0);
    if (kept !== inkedBlanks.length) {
      throw new PdfNativeTransactionError("unsupported-document", MRAI_NATIVE_GLYPHS_READ_AS_BLANKS_MESSAGE);
    }
  };
  keepsGlyphsReadAsBlanks(mixed);
  let extensionText = mixed
    ? mixed.filter((segment) => !segment.source).map((segment) => segment.text).join("")
    : wholeReplacementText;
  let mixedLayout: MixedShowLayout | null = null;
  if (mixed && !fixedFrame) {
    const metrics = await measurePdfNativeFontExtensionText({
      fontBytes: input.fontBytes,
      text: extensionText,
    });
    mixedLayout = mixedShowLayout(mixed, metrics.glyphs, fidelity.fontSize);
    if (mixedLayout === null && fidelity.readerChunk) {
      return invalid("Native font extension cannot set a reader chunk as a whole run.");
    }
    if (
      !fidelity.readerChunk &&
      (mixedLayout === null || mixedLayout.advance > singleRunCapacity + 0.001)
    ) {
      // The few substituted glyphs do not fit beside the source font's own
      // characters, but the whole run in the extension font may (a narrow
      // face standing in for a wide one); then the run is set whole.
      const whole = await measurePdfNativeFontExtensionText({
        fontBytes: input.fontBytes,
        text: wholeReplacementText,
      });
      if (
        mixedLayout === null ||
        (whole.widthAt1000 * fidelity.fontSize) / 1_000 <= singleRunCapacity + 0.001
      ) {
        mixed = null;
        mixedLayout = null;
        extensionText = wholeReplacementText;
        keepsGlyphsReadAsBlanks(mixed);
      }
    }
  }
  const fontPackage = reusable
    ? await buildMraiNativeFontAliasPackage({
        reusable,
        resourceName: name,
        firstObjectNumber,
        replacementText: input.replacementText,
      })
    : await buildPdfNativeFontExtensionPackage({
        fontBytes: input.fontBytes,
        text: extensionText,
        resourceName: name,
        firstObjectNumber,
      });
  let singleRunReplacementAdvance: number | null = null;
  let replacement: Uint8Array;
  let operatorSpan = fidelity.textOperatorSpan;
  let replacementLines: readonly string[] | null = null;
  let replacementLineAdvances: readonly number[] | null = null;
  const size = formatNumber(fidelity.fontSize);
  if (fixedFrame && linePlan) {
    const metrics = await measurePdfNativeFontExtensionText({
      fontBytes: input.fontBytes,
      text: extensionText,
    });
    if (
      fontPackage.encodedHex.length !== metrics.glyphs.length * 4 ||
      Math.abs(fontPackage.widthAt1000 - metrics.widthAt1000) > 0.000001
    ) {
      return invalid("Native font extension glyph metrics changed.");
    }
    const layout = linePlan.lines[0].layoutState!;
    if (
      linePlan.lines.some(
        (line) =>
          !line.layoutState ||
          Math.abs(line.layoutState.fontSize - layout.fontSize) > 0.000001 ||
          Math.abs(
            line.layoutState.characterSpacing - layout.characterSpacing,
          ) > 0.000001 ||
          Math.abs(line.layoutState.wordSpacing - layout.wordSpacing) >
            0.000001 ||
          Math.abs(line.layoutState.horizontalScale - layout.horizontalScale) >
            0.000001,
      ) ||
      Math.abs(layout.fontSize - fidelity.fontSize) > 0.000001
    ) {
      return invalid("Native font extension line layout state changed.");
    }
    const horizontalScale = layout.horizontalScale / 100;
    const characterSpacingAt1000 =
      (layout.characterSpacing * 1_000) / layout.fontSize;
    const encodings = new Map(
      (sourceProof.singleLineFit?.encodings ?? []).filter(
        ([character]) => [...character].length === 1,
      ),
    );
    const advances = new Map(sourceProof.singleLineFit?.advances ?? []);
    // Every unit the frame wraps by: a character the source font writes, or
    // one glyph of the extension font, each in the line's 1000-unit space.
    const effectiveGlyphs = mixed
      ? frameUnits(
          mixed,
          metrics.glyphs,
          (index) => fontPackage.encodedHex.slice(index * 4, index * 4 + 4),
          (character) => {
            const codes = encodings.get(character);
            const advance = advances.get(character);
            return codes === undefined || advance === undefined
              ? null
              : Object.freeze({
                  codes,
                  // The hint measures on the page, spacing and horizontal
                  // scale included; the frame counts in 1000 units.
                  advanceAt1000:
                    (advance * 1_000) / layout.fontSize / horizontalScale,
                });
          },
          characterSpacingAt1000,
          (advance) => (advance * 1_000) / layout.fontSize / horizontalScale,
        )
      : metrics.glyphs.map((glyph, index) =>
          Object.freeze({
            text: glyph.text,
            advanceAt1000: glyph.advanceAt1000 + characterSpacingAt1000,
            source: false,
            hex: fontPackage.encodedHex.slice(index * 4, index * 4 + 4),
          }),
        );
    if (effectiveGlyphs === null) {
      return invalid("Native font extension cannot lay a framed line out in both fonts.");
    }
    const groupStartIndexes = visualLineIndexes
      .map((lineIndex, index) =>
        index === 0 || lineIndex !== visualLineIndexes[index - 1] ? index : -1,
      )
      .filter((index) => index >= 0);
    let capacitiesAt1000 = groupStartIndexes.map(
      (index) =>
        ((linePlan.frameRight - linePlan.lines[index].textX) * 1_000) /
        layout.fontSize /
        horizontalScale,
    );
    // One line that ends at its frame's edge (a right-aligned date beside a
    // paper's title, the frame a Form's box) has no room to grow on the
    // right: it grows to the left instead, by the free space before it, and
    // its right end stays where it was -- the way a right-aligned paragraph
    // takes typing. The pen is moved back before the line's first glyph, so
    // every operator after it reads exactly as it did.
    let frameLeadShiftAt1000 = 0;
    if (groupStartIndexes.length === 1) {
      const totalAt1000 = effectiveGlyphs.reduce((total, glyph) => total + glyph.advanceAt1000, 0);
      const lastLine = linePlan.lines.at(-1)!;
      const originalAt1000 =
        ((lastLine.textX + lastLine.originalAdvance - linePlan.lines[0].textX) * 1_000) /
        layout.fontSize /
        horizontalScale;
      const shiftAt1000 = totalAt1000 - originalAt1000;
      const leftRoomAt1000 =
        ((sourceProof.singleLineFit?.leftRoom ?? 0) * 1_000) / layout.fontSize / horizontalScale;
      if (
        totalAt1000 > capacitiesAt1000[0] + 0.001 &&
        capacitiesAt1000[0] - originalAt1000 < RIGHT_ANCHORED_ROOM_AT_1000 &&
        shiftAt1000 > 0 &&
        shiftAt1000 <= leftRoomAt1000
      ) {
        frameLeadShiftAt1000 = shiftAt1000;
        capacitiesAt1000 = [totalAt1000 + 0.001];
      }
    }
    const wrappedGroups = wrapPdfNativeMeasuredGlyphsInFixedFrame({
      glyphs: effectiveGlyphs,
      capacitiesAt1000,
    });
    const wrapped = linePlan.lines.map((_, index) => {
      const visualLineIndex = visualLineIndexes[index];
      const group = wrappedGroups[visualLineIndex];
      return groupStartIndexes[visualLineIndex] === index
        ? group
        : Object.freeze({
            text: "",
            glyphStart: group.glyphEnd,
            glyphEnd: group.glyphEnd,
            advanceAt1000: 0,
          });
    });
    const parts: Uint8Array[] = [];
    const first = linePlan.lines[0];
    const last = linePlan.lines.at(-1)!;
    operatorSpan = {
      start: first.operatorSpan.start,
      end: last.operatorSpan.end,
      sha256: await sha256PdfBytes(
        stream.streamBytes.slice(
          first.operatorSpan.start,
          last.operatorSpan.end,
        ),
      ),
    };
    let sourceCursor = operatorSpan.start;
    for (const [index, line] of linePlan.lines.entries()) {
      if (
        line.operatorSpan.start < sourceCursor ||
        line.operatorSpan.end > operatorSpan.end
      ) {
        return invalid("Native font extension line spans overlap.");
      }
      parts.push(
        stream.streamBytes.slice(sourceCursor, line.operatorSpan.start),
      );
      const wrappedLine = wrapped[index];
      const linePrefix =
        line.operatorKind === "'" || line.operatorKind === "'-literal"
          ? "T* "
          : "";
      if (!wrappedLine.text) {
        parts.push(ascii(`${linePrefix}<> Tj`));
      } else {
        const restoreName = line.resourceName ?? fidelity.resourceName;
        // Consecutive units of one font are one show; the line returns to
        // its own font after every extension show, so the operators after
        // it read exactly as they did.
        const shows: string[] = [];
        let cursor = wrappedLine.glyphStart;
        while (cursor < wrappedLine.glyphEnd) {
          const { source } = effectiveGlyphs[cursor];
          let end = cursor;
          let hex = "";
          while (end < wrappedLine.glyphEnd && effectiveGlyphs[end].source === source) {
            hex += effectiveGlyphs[end].hex;
            end += 1;
          }
          shows.push(
            source
              ? `<${hex}> Tj`
              : `/${name} ${size} Tf <${hex}> Tj /${restoreName} ${size} Tf`,
          );
          cursor = end;
        }
        const lead = index === 0 && frameLeadShiftAt1000 > 0
          ? `[${formatLeadOffset(frameLeadShiftAt1000)}] TJ `
          : "";
        parts.push(ascii(`${linePrefix}${lead}${shows.join(" ")}`));
      }
      sourceCursor = line.operatorSpan.end;
    }
    replacement = concat([
      ...parts,
      stream.streamBytes.slice(sourceCursor, operatorSpan.end),
    ]);
    replacementLines = Object.freeze(wrapped.map((line) => line.text));
    replacementLineAdvances = Object.freeze(
      wrapped.map(
        (line) =>
          (line.advanceAt1000 * layout.fontSize * horizontalScale) / 1_000,
      ),
    );
  } else {
    let replacementAdvance: number;
    let show: Uint8Array;
    if (mixed && mixedLayout) {
      if (
        fontPackage.encodedHex.length !== mixedLayout.glyphCount * 4 ||
        Math.abs(fontPackage.widthAt1000 - mixedLayout.widthAt1000) > 0.000001
      ) {
        return invalid("Native font extension glyph metrics changed.");
      }
      const parts = mixed.map((segment, index) => {
        const range = mixedLayout!.ranges[index];
        return range === null
          ? `<${hexOf(segment.codes)}> Tj`
          : `/${name} ${size} Tf <${fontPackage.encodedHex.slice(range[0] * 4, range[1] * 4)}> Tj /${fidelity.resourceName} ${size} Tf`;
      });
      replacementAdvance = mixedLayout.advance;
      show = ascii(parts.join(" "));
    } else {
      replacementAdvance = (fontPackage.widthAt1000 * fidelity.fontSize) / 1_000;
      show = ascii(
        `/${name} ${size} Tf <${fontPackage.encodedHex}> Tj /${fidelity.resourceName} ${size} Tf`,
      );
    }
    // A run that ends at its frame's edge (nothing but a fraction of an em
    // beside it) and that the glyphs do not fit grows to the left, by the
    // free space before it, keeping its right end: the pen is moved back
    // before its first glyph and ends where the run ended, so the line's
    // later operators read exactly as they did.
    const runHorizontalScale = (fidelity.sameLineLayoutState?.horizontalScale ?? 100) / 100;
    const leftShift = replacementAdvance - fidelity.originalAdvance;
    if (
      replacementAdvance > singleRunCapacity + 0.001 &&
      !fidelity.lineShift &&
      !fidelity.readerChunk &&
      singleRunCapacity - fidelity.originalAdvance <
        (RIGHT_ANCHORED_ROOM_AT_1000 * fidelity.fontSize) / 1_000 &&
      leftShift > 0 &&
      leftShift <= (sourceProof.singleLineFit?.leftRoom ?? 0)
    ) {
      show = concat([
        ascii(`[${formatLeadOffset((leftShift * 1_000) / fidelity.fontSize / runHorizontalScale)}] TJ `),
        show,
      ]);
      replacementAdvance = fidelity.originalAdvance;
    }
    singleRunReplacementAdvance = replacementAdvance;
    // The installed font is usually wider than the one it stands in for. A
    // single line may grow into the free page space beside it that the
    // source proof measured; replay re-derives that space and refuses the
    // record when it is not exactly that.
    if (replacementAdvance > singleRunCapacity + 0.001) {
      // The free space beside the run could not hold the glyphs; a proof
      // that moves the line's later runs may leave more room.
      if (!fidelity.lineShift && !input.preferLineShift) {
        return planPdfNativeFontExtensionRunTextEdit({ ...input, preferLineShift: true });
      }
      return invalid("Native font extension replacement exceeds its text run.");
    }
    // A proof that moves the line's later runs spans more than the show:
    // the tail is rewritten by this font's own width change. A run group's
    // show is its first run; the group's later runs, emptied, come between.
    replacement = concat([
      show,
      ...(runGroupRest ? [runGroupRest] : []),
      ...(fidelity.lineShift
        ? [
            shiftedLineTail(
              LATIN1_TEXT.decode(stream.streamBytes),
              fidelity.lineShift.showEnd,
              { envelopeEnd: operatorSpan.end, operators: fidelity.lineShift.operators },
              replacementAdvance - fidelity.originalAdvance,
            ),
          ]
        : []),
    ]);
  }
  // A line the page draws again in its place is set in every copy.
  const duplicateSpans = fontSetDuplicateSpans(
    sourceProof,
    operatorSpan,
    !fixedFrame && runGroupRest === null && !fidelity.lineShift,
  );
  const grownSourceFidelity =
    !fixedFrame &&
    !fidelity.lineShift &&
    (singleRunReplacementAdvance ??
      (fontPackage.widthAt1000 * fidelity.fontSize) / 1_000) >
      fidelity.originalAdvance + 0.001
      ? { ...fidelity, sameLineCapacity: singleRunCapacity }
      : fidelity;
  const pageFontExtensionMutation =
    await planMraiNativePageFontExtensionMutation({
      sourceBytes,
      pageNumber: page.pageNumber,
      fontPackage,
      reservedObjects: consolidation ? 1 : isolation ? isolation.chain.length : 0,
      ...(formTarget
        ? {
            form: {
              objectNumber: formTarget.objectNumber,
              generation: formTarget.generation,
            },
          }
        : {}),
    });
  const resourceOwner = formTarget
    ? null
    : await inspectMraiNativeObject(sourceBytes, page.resourceAuthority.owner);
  const prepared = await prepareMraiNativeContentStreamTransaction(
    sourceBytes,
    await withStructActualText(sourceBytes, page.pageNumber, {
      identity: stream.identity,
      expectedObjectSha256: stream.objectSha256,
      expectedStreamSha256: stream.streamSha256,
      operatorSpan: {
        start: operatorSpan.start,
        end: operatorSpan.end,
        expectedSha256: operatorSpan.sha256,
      },
      replacement,
      ...(duplicateSpans ? { duplicateSpans } : {}),
      untouchedObjects: withProvenResourcePreconditions(
        resourceOwner
          ? [
              {
                objectNumber: resourceOwner.identity.objectNumber,
                generation: resourceOwner.identity.generation,
                expectedObjectSha256: resourceOwner.objectSha256,
                role: "resource",
              },
            ]
          : [],
        sourceProof,
      ),
      pageFontExtensionMutation,
      ...(consolidation ? { pageContentConsolidation: consolidation } : {}),
      ...(isolation ? { formInstanceIsolation: isolation } : {}),
    }, [fontPackage], { before: input.selection.item.str, after: input.replacementText }),
  );
  await verifyMraiNativeContentStreamCandidate(sourceBytes, prepared);
  const record = JSON.parse(
    JSON.stringify({
      id: input.transactionId,
      schemaVersion: MRAI_NATIVE_FONT_EXTENSION_TEXT_SCHEMA_VERSION,
      transactionSchemaVersion: prepared.schemaVersion,
      kind: "font-extension-text-replacement",
      featureId: PDF_NATIVE_FONT_EXTENSION_TEXT_FEATURE_ID,
      controlId: PDF_NATIVE_FONT_EXTENSION_TEXT_CONTROL_ID,
      action: MRAI_NATIVE_FONT_EXTENSION_TEXT_ACTION,
      transactionId: input.transactionId,
      baseSessionRevision: input.baseSessionRevision,
      selectionRevision: input.selection.selectionRevision,
      source: {
        resourceId: input.source.resourceId,
        resourceVersionId: input.source.resourceVersionId,
        sha256: sourceSha256,
      },
      target: {
        objectNumber: stream.identity.objectNumber,
        generation: stream.identity.generation,
        expectedObjectSha256: stream.objectSha256,
        expectedStreamSha256: stream.streamSha256,
      },
      operatorSpan: prepared.request.operatorSpan,
      replacement: [...replacement],
      // The copies of the show the page draws again in its place.
      ...(duplicateSpans ? { duplicateSpans } : {}),
      replacementSha256: prepared.replacementSha256,
      ...pdfNativeStructActualTextRecordField(prepared.request),
      ...pdfNativePdfaIdentificationRecordField(prepared),
      untouchedObjects: prepared.request.untouchedObjects,
      formInstanceIsolation: isolation,
      pageExtGStateMutation: null,
      pageContentConsolidation: consolidation,
      pageFontExtensionMutation: {
        ...pageFontExtensionMutation,
        fontPackage: serializeFontMutationPackage(fontPackage),
      },
      derivedCandidate: {
        sha256: prepared.candidateSha256,
        byteLength: prepared.candidateBytes.length,
        sourceLength: prepared.sourceLength,
        sourceXrefOffset: prepared.sourceXrefOffset,
        changedObjectOffset: prepared.changedObjectOffset,
        incrementalXrefOffset: prepared.incrementalXrefOffset,
        changedObjectSha256: prepared.changedObjectSha256,
        changedStreamSha256: prepared.changedStreamSha256,
        revisedObjects: derivedObjects(prepared),
      },
      listFidelity: null,
      textFidelity: null,
      embeddedImageRemoval: null,
      embeddedImageReplacement: null,
      outlineMutation: null,
      vectorGeometry: null,
      vectorDeletion: null,
      vectorStyle: null,
      fontExtensionText: {
        schemaVersion: MRAI_NATIVE_FONT_EXTENSION_TEXT_SCHEMA_VERSION,
        sourceFidelity: grownSourceFidelity,
        // The whole show's text: a reader chunk's replacement with the
        // codes around it, the way the show reads after the edit.
        replacementText: wholeReplacementText,
        ...(mixed ? { extensionText } : {}),
        ...(replacementLines
          ? {
              replacementLines,
              replacementLineAdvances,
              sourceProofReplacement: [...sourceProof.request.replacement],
            }
          : {}),
      },
    }),
  ) as PdfSharedStateRecord;
  const plan: PdfNativeFontExtensionTextPlan = Object.freeze({ sourceProof, prepared, record });
  if (singleRunReplacementAdvance !== null) {
    singleRunReplacementAdvances.set(plan, singleRunReplacementAdvance);
  }
  return plan;
}

type PdfNativeFontExtensionTextEditInput = Omit<
  Parameters<typeof planPdfNativeFontExtensionRunTextEdit>[0],
  "groupRunTried"
>;

/**
 * A text edit whose new text holds characters the run's font lacks. The edit
 * is written, the way a desktop editor writes it, with the missing
 * characters drawn in a substitute face the document gains:
 *
 * 1. One exact source run or fixed frame keeps every character its own font
 *    writes and sets only the others in the installed face (the run route).
 * 2. Any other paragraph -- several lines, several runs, several fonts -- is
 *    rewritten by the paragraph tier: each character stays in its own font,
 *    one its font lacks goes to another font of the paragraph that has it,
 *    and only one no font of the paragraph has is set in the installed face,
 *    at the size, spacing and colour of the character it stands beside.
 *
 * No refusal of the run route is final while the paragraph route can write
 * the edit; when neither can, the run route's reason is the one reported.
 */
export async function planPdfNativeFontExtensionTextEdit(
  input: PdfNativeFontExtensionTextEditInput,
): Promise<PdfNativeFontExtensionTextPlan> {
  // A run whose face is bold: the substitute face has no bold, so the run
  // route (the installed face in place of the show) would set the text
  // regular; the paragraph route draws it stroked as a synthetic bold
  
  // Where that route cannot set it, the run route still writes it. A Latin
  // character or two the face leaves out (a 'V' missing from a subset's
  // widths) stays with the run route, which keeps the line's own objects.
  let boldFace = false;
  if (Array.from(input.replacementText).some((character) =>
    /\p{L}/u.test(character) && !/\p{Script=Latin}/u.test(character))) {
    try {
      const sourceBytes = ownedPdfSourceBytes(input.sourceBytes);
      const proof = await proveMraiNativeListTextSelectionForWriter({ sourceBytes, selection: input.selection });
      boldFace = (await readMraiNativeEvidenceFaceStyle(sourceBytes, proof.evidence)).bold;
    } catch {
      // The routes below say why the run cannot be proved.
    }
  }
  // Formats with the text: the paragraph route sets them, the run route
  // cannot (a part made bold as the box's new words, missing from its
  
  if (input.paragraphFormat || input.paragraphFrame) return planPdfNativeFontExtensionParagraphTextEdit(input);
  // Words typed where two rows meet, the editor saying which row holds them:
  // the paragraph route first, where its rows settle a line break (the run
  // route sets them in the run it finds, whichever row that is). Where they
  // settle none, or it cannot, the routes are tried as ever.
  if (input.rowStarts?.length) {
    try {
      return await planPdfNativeFontExtensionParagraphTextEdit(input, true);
    } catch (rowError) {
      if (rowError instanceof Error && /source is stale/u.test(rowError.message)) throw rowError;
    }
  }
  if (boldFace) {
    try {
      return await planPdfNativeFontExtensionParagraphTextEdit(input);
    } catch (paragraphError) {
      if (paragraphError instanceof Error && /source is stale/u.test(paragraphError.message)) throw paragraphError;
    }
  }
  try {
    return await planPdfNativeFontExtensionRunTextEdit(input);
  } catch (runError) {
    if (runError instanceof Error && /source is stale|Invalid native font extension request/u.test(runError.message)) {
      throw runError;
    }
    try {
      return await planPdfNativeFontExtensionParagraphTextEdit(input);
    } catch (paragraphError) {
      const reason = (error: unknown) => (error instanceof Error ? error.message : String(error));
      // Neither route can set the text in place: a refusal of the document's
      
      throw new PdfNativeTransactionError("unsupported-document",
        // A glyph even the substitute face lacks is not the missing-glyph
        // refusal the font extension answers any more; it is said plainly.
        `${reason(runError)} | Paragraph with a substitute face: ${reason(paragraphError)}`
          .replaceAll(MRAI_NATIVE_LIST_TEXT_MISSING_GLYPH_MESSAGE, "The substitute face has no glyph for a character of the text.")
          .slice(0, 4_000));
    }
  }
}

/** The paragraph route: the paragraph rewrite with a substitute face. */
async function planPdfNativeFontExtensionParagraphTextEdit(
  input: PdfNativeFontExtensionTextEditInput,
  /** Refused unless the editor's rows settle a line break of the edit. */
  rowsMustDecide = false,
): Promise<PdfNativeFontExtensionTextPlan> {
  const sourceBytes = ownedPdfSourceBytes(input.sourceBytes);
  const sourceSha256 = await sha256PdfBytes(sourceBytes);
  if (sourceSha256 !== input.selection.sourceSha256.toLowerCase()) {
    return invalid("Native font extension source is stale.");
  }
  const name = await resourceName({
    sourceSha256,
    pageNumber: input.selection.pageNumber,
    targetObjectNumber: 0,
    operatorStart: 0,
    replacementText: input.replacementText,
  });
  // The face goes in as a subset of the glyphs the paragraph draws in it:
  // the rewrite's block is a text object of its own, so the subset's own
  // glyph numbering is all it ever names (the server proves each glyph's
  // outline against the canonical face, whatever its number).
  const nextObjectNumber = nextMraiNativeIndirectObjectNumber(sourceBytes);
  const subsetPackage = (text: string, firstObjectNumber: number) => buildPdfNativeFontExtensionPackage({
    fontBytes: input.fontBytes, text, resourceName: name, firstObjectNumber, subset: true,
  });
  let installed: Readonly<{ text: string; encodedHex: string }> | null = null;
  const plan = await planMraiNativeParagraphRewriteWithSubstitute({
    sourceBytes,
    selection: input.selection,
    replacementText: input.replacementText,
    ...(input.paragraphFormat ? { format: input.paragraphFormat } : {}),
    ...(input.paragraphFrame ? { frame: input.paragraphFrame } : {}),
    ...(input.rowStarts?.length && !input.paragraphFrame
      ? { rowEdit: { rowStarts: input.rowStarts, ...(rowsMustDecide ? { mustDecide: true } : {}) } } : {}),
    substituteFaceName: PDF_NATIVE_FALLBACK_FONT_NAME,
    async install(text) {
      let metrics: Awaited<ReturnType<typeof measurePdfNativeFontExtensionText>>;
      let provisional: Awaited<ReturnType<typeof buildPdfNativeFontExtensionPackage>>;
      try {
        metrics = await measurePdfNativeFontExtensionText({ fontBytes: input.fontBytes, text });
        provisional = await subsetPackage(text, nextObjectNumber);
      } catch {
        // The substitute face lacks the character too.
        return null;
      }
      // One glyph for each character: the face is laid out cmap-exact, so a
      // character it has is one glyph of it.
      const wanted = [...text];
      if (
        metrics.glyphs.length !== wanted.length ||
        metrics.glyphs.some((glyph, index) => glyph.text !== wanted[index]) ||
        provisional.encodedHex.length !== wanted.length * 4
      ) {
        return null;
      }
      installed = Object.freeze({ text, encodedHex: provisional.encodedHex });
      return Object.freeze({
        resourceName: name,
        glyphs: metrics.glyphs.map((glyph, index) => Object.freeze({
          cid: Number.parseInt(provisional.encodedHex.slice(index * 4, index * 4 + 4), 16),
          text: glyph.text,
          advanceAt1000: glyph.advanceAt1000,
        })),
      });
    },
  });
  const substitute = plan.substitute;
  const glyphs = installed as Readonly<{ text: string; encodedHex: string }> | null;
  if (!substitute || !glyphs || substitute.resourceName !== name || substitute.text !== glyphs.text) {
    return invalid("Native font extension paragraph sets no character in the substitute face.");
  }
  const request = plan.request;
  const consolidation = request.pageContentConsolidation ?? null;
  if (consolidation && consolidation.stream.objectNumber !== nextObjectNumber) {
    return invalid("Consolidated content stream allocation changed.");
  }
  const fontPackage = await subsetPackage(glyphs.text, nextObjectNumber + (consolidation ? 1 : 0));
  // The block drew each substituted character by its code in the subset;
  // the installed program must name the same glyphs.
  if (fontPackage.encodedHex !== glyphs.encodedHex) {
    return invalid("Native font extension glyph metrics changed.");
  }
  const page = await inspectMraiNativePageContentGraph(sourceBytes, input.selection.pageNumber);
  const targetIsPageStream = page.contentStreams.some(
    (candidate) =>
      candidate.objectNumber === request.identity.objectNumber &&
      candidate.generation === request.identity.generation,
  );
  // Text a Form XObject draws takes the face into that form's own resources.
  const formTarget = !consolidation && !targetIsPageStream && !plan.pageContent ? request.identity : null;
  if (!consolidation && !targetIsPageStream && !formTarget) {
    return invalid("Native font extension stream ownership is ambiguous.");
  }
  const pageFontExtensionMutation = await planMraiNativePageFontExtensionMutation({
    sourceBytes,
    pageNumber: page.pageNumber,
    fontPackage,
    reservedObjects: consolidation ? 1 : 0,
    ...(formTarget
      ? { form: { objectNumber: formTarget.objectNumber, generation: formTarget.generation } }
      : {}),
  });
  const prepared = await prepareMraiNativeContentStreamTransaction(
    sourceBytes,
    await withStructActualText(sourceBytes, page.pageNumber, { ...request, pageFontExtensionMutation }, [fontPackage], { before: input.selection.item.str, after: input.replacementText }),
  );
  await verifyMraiNativeContentStreamCandidate(sourceBytes, prepared);
  const record = JSON.parse(
    JSON.stringify({
      id: input.transactionId,
      schemaVersion: MRAI_NATIVE_FONT_EXTENSION_TEXT_SCHEMA_VERSION,
      transactionSchemaVersion: prepared.schemaVersion,
      kind: "font-extension-text-replacement",
      featureId: PDF_NATIVE_FONT_EXTENSION_TEXT_FEATURE_ID,
      controlId: PDF_NATIVE_FONT_EXTENSION_TEXT_CONTROL_ID,
      action: MRAI_NATIVE_FONT_EXTENSION_TEXT_ACTION,
      transactionId: input.transactionId,
      baseSessionRevision: input.baseSessionRevision,
      selectionRevision: input.selection.selectionRevision,
      source: {
        resourceId: input.source.resourceId,
        resourceVersionId: input.source.resourceVersionId,
        sha256: sourceSha256,
      },
      target: {
        objectNumber: request.identity.objectNumber,
        generation: request.identity.generation,
        expectedObjectSha256: request.expectedObjectSha256,
        expectedStreamSha256: request.expectedStreamSha256,
      },
      operatorSpan: prepared.request.operatorSpan,
      replacement: [...request.replacement],
      replacementSha256: prepared.replacementSha256,
      ...pdfNativeStructActualTextRecordField(prepared.request),
      ...pdfNativePdfaIdentificationRecordField(prepared),
      untouchedObjects: prepared.request.untouchedObjects,
      formInstanceIsolation: null,
      pageExtGStateMutation: null,
      pageContentConsolidation: consolidation,
      pageFontExtensionMutation: {
        ...pageFontExtensionMutation,
        fontPackage: serializeFontMutationPackage(fontPackage),
      },
      derivedCandidate: {
        sha256: prepared.candidateSha256,
        byteLength: prepared.candidateBytes.length,
        sourceLength: prepared.sourceLength,
        sourceXrefOffset: prepared.sourceXrefOffset,
        changedObjectOffset: prepared.changedObjectOffset,
        incrementalXrefOffset: prepared.incrementalXrefOffset,
        changedObjectSha256: prepared.changedObjectSha256,
        changedStreamSha256: prepared.changedStreamSha256,
        revisedObjects: derivedObjects(prepared),
      },
      listFidelity: null,
      textFidelity: null,
      embeddedImageRemoval: null,
      embeddedImageReplacement: null,
      outlineMutation: null,
      vectorGeometry: null,
      vectorDeletion: null,
      vectorStyle: null,
      fontExtensionText: {
        schemaVersion: MRAI_NATIVE_FONT_EXTENSION_TEXT_SCHEMA_VERSION,
        // The paragraph as the rewrite proved and set it; the characters it
        // set in the installed face are the package's text.
        sourceFidelity: plan.evidence,
        replacementText: input.replacementText,
        extensionText: glyphs.text,
      },
    }),
  ) as PdfSharedStateRecord;
  const sourceProof: PdfNativeListTextSelectionProof = Object.freeze({
    target: Object.freeze({ objectNumber: request.identity.objectNumber, generation: request.identity.generation }),
    evidence: plan.evidence,
    properties: null,
  });
  return Object.freeze({ sourceProof, prepared, record });
}

/**
 * A property edit that sets part of a run -- or all of it -- in a family the
 * PDF does not hold (WPS: two characters of a line set in Arial): the faces
 * MRAI serves for the family are installed as subsets, one font per slice of
 * the face the text needs (native-embedded-fonts.ts), and the run is written
 * as its parts exactly as any part format writes it, each part in its own
 * font. The write is a font-extension record whose source fidelity is the
 * property edit itself: the server proves every installed glyph against the
 * pinned slice and replays the parts with the installed faces' metrics, and
 * Undo takes the fonts off the page with the edit (font-extension history).
 */
export async function planPdfNativeEmbeddedFontPropertyEdit(input: {
  readonly sourceBytes: Uint8Array;
  readonly source: PdfNativeFontExtensionSourceIdentity;
  readonly selection: PdfNativeListTextSelectionEvidence;
  readonly replacementText: string;
  readonly proof: PdfNativeListTextPropertyReplacementInput["proof"];
  readonly properties: PdfNativeListTextPropertyReplacementInput["properties"];
  readonly keepTextRules?: PdfNativeListTextPropertyReplacementInput["keepTextRules"];
  readonly previousTurn?: PdfNativeListTextPropertyReplacementInput["previousTurn"];
  readonly library: PdfNativeEmbeddedFontLibrary;
  readonly transactionId: string;
  readonly baseSessionRevision: number;
}): Promise<PdfNativeFontExtensionTextPlan> {
  if (
    !input.source.resourceId ||
    !input.source.resourceVersionId ||
    !input.transactionId ||
    !Number.isSafeInteger(input.baseSessionRevision) ||
    input.baseSessionRevision < 0
  ) {
    return invalid("Invalid native embedded font request.");
  }
  const sourceBytes = ownedPdfSourceBytes(input.sourceBytes);
  const sourceSha256 = await sha256PdfBytes(sourceBytes);
  if (sourceSha256 !== input.selection.sourceSha256.toLowerCase()) {
    return invalid("Native embedded font source is stale.");
  }
  // The editor's preview of a format set in a served face and the write
  // when the line is left ask this the very same question: planned once
  // (embeddedPropertyPlans), each asker given its own copy under its own
  // transaction (the id is the record's alone: the bytes do not hold it).
  const key = embeddedPropertyPlanKey(input, sourceSha256);
  if (key !== null) {
    const now = Date.now();
    for (const [kept, entry] of embeddedPropertyPlans) if (now - entry.at > EMBEDDED_PROPERTY_PLAN_KEPT_MS) embeddedPropertyPlans.delete(kept);
    let entry = embeddedPropertyPlans.get(key);
    if (!entry) {
      const planned = planPdfNativeEmbeddedFontPropertyEditOnce(input, sourceBytes, sourceSha256);
      const settled = { at: now, plan: planned };
      entry = settled;
      embeddedPropertyPlans.set(key, settled);
      planned.catch(() => { if (embeddedPropertyPlans.get(key) === settled) embeddedPropertyPlans.delete(key); });
      while (embeddedPropertyPlans.size > EMBEDDED_PROPERTY_PLANS_KEPT) {
        embeddedPropertyPlans.delete(embeddedPropertyPlans.keys().next().value!);
      }
    }
    const answer = await entry.plan;
    try {
      const copy = structuredClone(answer);
      return Object.freeze({ ...copy, record: { ...copy.record, id: input.transactionId,
        transactionId: input.transactionId, baseSessionRevision: input.baseSessionRevision } as PdfSharedStateRecord });
    } catch {
      embeddedPropertyPlans.delete(key);
    }
  }
  return planPdfNativeEmbeddedFontPropertyEditOnce(input, sourceBytes, sourceSha256);
}

const EMBEDDED_PROPERTY_PLANS_KEPT = 4;
const EMBEDDED_PROPERTY_PLAN_KEPT_MS = 60_000;
const embeddedPropertyPlans = new Map<string, { at: number;
  plan: Promise<Awaited<ReturnType<typeof planPdfNativeEmbeddedFontPropertyEditOnce>>> }>();

/** Every input but the transaction (and the bytes, by their digest); null
 * when the library cannot say what it holds. */
function embeddedPropertyPlanKey(
  input: Parameters<typeof planPdfNativeEmbeddedFontPropertyEdit>[0],
  sourceSha256: string,
): string | null {
  if (!input.library.fingerprint) return null;
  const { sourceBytes: _bytes, library, transactionId: _id, baseSessionRevision: _revision, ...question } = input;
  void _bytes; void _id; void _revision;
  try {
    return JSON.stringify([sourceSha256, library.fingerprint, question]);
  } catch {
    return null;
  }
}

async function planPdfNativeEmbeddedFontPropertyEditOnce(
  input: Parameters<typeof planPdfNativeEmbeddedFontPropertyEdit>[0],
  sourceBytes: Uint8Array,
  sourceSha256: string,
) {
  const installation = createPdfNativeEmbeddedFontInstallation({
    library: input.library,
    sourceSha256,
    pageNumber: input.selection.pageNumber,
  });
  const plan = await planMraiNativeListTextPropertyReplacement({
    sourceBytes,
    selection: input.selection,
    replacementText: input.replacementText,
    proof: input.proof,
    properties: input.properties,
    ...(input.keepTextRules?.length ? { keepTextRules: input.keepTextRules } : {}),
    ...(input.previousTurn ? { previousTurn: input.previousTurn } : {}),
    embeddedFonts: installation,
  });
  const packages = await installation.packages();
  if (packages.length === 0) {
    return invalid("Native embedded font edit sets no text in an installed face.");
  }
  const request = plan.request;
  const consolidation = request.pageContentConsolidation ?? null;
  const page = await inspectMraiNativePageContentGraph(sourceBytes, input.selection.pageNumber);
  const targetIsPageStream = page.contentStreams.some(
    (candidate) =>
      candidate.objectNumber === request.identity.objectNumber &&
      candidate.generation === request.identity.generation,
  );
  // A run a Form XObject draws keeps its fonts in the form's resources,
  // which this write does not install into yet.
  if (request.formInstanceIsolation || (!consolidation && !targetIsPageStream)) {
    throw new PdfNativeEmbeddedFontError("unsupported-place", "");
  }
  const pageFontExtensionMutation = await planMraiNativePageFontExtensionMutation({
    sourceBytes,
    pageNumber: page.pageNumber,
    fontPackage: packages[0],
    ...(packages.length > 1 ? { additionalFontPackages: packages.slice(1) } : {}),
    reservedObjects: consolidation ? 1 : 0,
  });
  const prepared = await prepareMraiNativeContentStreamTransaction(
    sourceBytes,
    await withStructActualText(sourceBytes, page.pageNumber, { ...request, pageFontExtensionMutation }, packages, { before: input.selection.item.str, after: input.replacementText }),
  );
  await verifyMraiNativeContentStreamCandidate(sourceBytes, prepared);
  const record = JSON.parse(
    JSON.stringify({
      id: input.transactionId,
      schemaVersion: MRAI_NATIVE_FONT_EXTENSION_TEXT_SCHEMA_VERSION,
      transactionSchemaVersion: prepared.schemaVersion,
      kind: "font-extension-text-replacement",
      featureId: PDF_NATIVE_FONT_EXTENSION_TEXT_FEATURE_ID,
      controlId: PDF_NATIVE_FONT_EXTENSION_TEXT_CONTROL_ID,
      action: MRAI_NATIVE_FONT_EXTENSION_TEXT_ACTION,
      transactionId: input.transactionId,
      baseSessionRevision: input.baseSessionRevision,
      selectionRevision: input.selection.selectionRevision,
      source: {
        resourceId: input.source.resourceId,
        resourceVersionId: input.source.resourceVersionId,
        sha256: sourceSha256,
      },
      target: {
        objectNumber: request.identity.objectNumber,
        generation: request.identity.generation,
        expectedObjectSha256: request.expectedObjectSha256,
        expectedStreamSha256: request.expectedStreamSha256,
      },
      operatorSpan: prepared.request.operatorSpan,
      replacement: [...request.replacement],
      replacementSha256: prepared.replacementSha256,
      ...(prepared.request.textRules ? { textRules: prepared.request.textRules } : {}),
      ...pdfNativeStructActualTextRecordField(prepared.request),
      ...pdfNativePdfaIdentificationRecordField(prepared),
      untouchedObjects: prepared.request.untouchedObjects,
      formInstanceIsolation: null,
      pageExtGStateMutation: null,
      pageContentConsolidation: consolidation,
      pageFontExtensionMutation: {
        ...pageFontExtensionMutation,
        fontPackage: serializeFontMutationPackage(pageFontExtensionMutation.fontPackage),
        ...(pageFontExtensionMutation.additionalFontPackages
          ? {
              additionalFontPackages: pageFontExtensionMutation.additionalFontPackages.map(
                serializePdfNativeFontExtensionPackage,
              ),
            }
          : {}),
      },
      derivedCandidate: {
        sha256: prepared.candidateSha256,
        byteLength: prepared.candidateBytes.length,
        sourceLength: prepared.sourceLength,
        sourceXrefOffset: prepared.sourceXrefOffset,
        changedObjectOffset: prepared.changedObjectOffset,
        incrementalXrefOffset: prepared.incrementalXrefOffset,
        changedObjectSha256: prepared.changedObjectSha256,
        changedStreamSha256: prepared.changedStreamSha256,
        revisedObjects: derivedObjects(prepared),
      },
      listFidelity: null,
      textFidelity: null,
      embeddedImageRemoval: null,
      embeddedImageReplacement: null,
      outlineMutation: null,
      vectorGeometry: null,
      vectorDeletion: null,
      vectorStyle: null,
      fontExtensionText: {
        schemaVersion: MRAI_NATIVE_FONT_EXTENSION_TEXT_SCHEMA_VERSION,
        // The run's property edit as planned and proved: its parts name the
        // installed fonts, whose packages are the mutation's.
        sourceFidelity: plan.evidence,
        replacementText: plan.evidence.replacementDisplayText,
      },
    }),
  ) as PdfSharedStateRecord;
  const sourceProof: PdfNativeListTextSelectionProof = Object.freeze({
    target: Object.freeze({ objectNumber: request.identity.objectNumber, generation: request.identity.generation }),
    evidence: plan.evidence,
    properties: plan.properties ?? null,
  });
  return Object.freeze({ sourceProof, prepared, record });
}

/**
 * A paragraph set again with formats (native list text: a line drawn as
 * several runs, or a run a format no longer fits) where a format names a
 * family the PDF does not hold: the faces are installed as subsets exactly as
 * a property edit installs them, and the paragraph rewrite sets each
 * character the family has in them, every other in its own face. The record
 * is a font-extension record whose source fidelity is the rewrite.
 */
export async function planPdfNativeEmbeddedFontParagraphFormatEdit(input: {
  readonly sourceBytes: Uint8Array;
  readonly source: PdfNativeFontExtensionSourceIdentity;
  readonly selection: PdfNativeListTextSelectionEvidence;
  readonly replacementText: string;
  readonly format: PdfNativeParagraphFormat;
  readonly fontOptions?: readonly PdfNativeTextFontOption[];
  readonly library: PdfNativeEmbeddedFontLibrary;
  readonly transactionId: string;
  readonly baseSessionRevision: number;
  /** Refused, not written with a notice, where it runs into other text. */
  readonly refuseOverOtherText?: boolean;
}): Promise<PdfNativeFontExtensionTextPlan> {
  if (!input.source.resourceId || !input.source.resourceVersionId || !input.transactionId ||
    !Number.isSafeInteger(input.baseSessionRevision) || input.baseSessionRevision < 0) {
    return invalid("Invalid native embedded font request.");
  }
  const sourceBytes = ownedPdfSourceBytes(input.sourceBytes);
  const sourceSha256 = await sha256PdfBytes(sourceBytes);
  if (sourceSha256 !== input.selection.sourceSha256.toLowerCase()) {
    return invalid("Native embedded font source is stale.");
  }
  const installation = createPdfNativeEmbeddedFontInstallation({
    library: input.library,
    sourceSha256,
    pageNumber: input.selection.pageNumber,
  });
  const plan = pdfNativeParagraphPlanClearOfOtherText(await planMraiNativeListTextParagraphFormat({
    sourceBytes,
    selection: input.selection,
    replacementText: input.replacementText,
    format: input.format,
    ...(input.fontOptions ? { fontOptions: input.fontOptions } : {}),
    embeddedFonts: installation,
  }), input.refuseOverOtherText === true);
  const packages = await installation.packages();
  const request = plan.request;
  const consolidation = request.pageContentConsolidation ?? null;
  const page = await inspectMraiNativePageContentGraph(sourceBytes, input.selection.pageNumber);
  const targetIsPageStream = page.contentStreams.some((candidate) =>
    candidate.objectNumber === request.identity.objectNumber && candidate.generation === request.identity.generation);
  if (request.formInstanceIsolation || (!consolidation && !targetIsPageStream)) {
    throw new PdfNativeEmbeddedFontError("unsupported-place", "");
  }
  // No character of the text is one the family has (a Latin face asked of a
  // Chinese line): nothing is installed, the formats' other fields are the
  // paragraph's own write.
  const pageFontExtensionMutation = packages.length
    ? await planMraiNativePageFontExtensionMutation({
      sourceBytes,
      pageNumber: page.pageNumber,
      fontPackage: packages[0],
      ...(packages.length > 1 ? { additionalFontPackages: packages.slice(1) } : {}),
      reservedObjects: consolidation ? 1 : 0,
    })
    : null;
  if (!pageFontExtensionMutation) {
    return invalid("Native embedded font paragraph sets no character in an installed face.");
  }
  const prepared = await prepareMraiNativeContentStreamTransaction(
    sourceBytes,
    await withStructActualText(sourceBytes, page.pageNumber, { ...request, pageFontExtensionMutation }, packages,
      { before: input.selection.item.str, after: input.replacementText }),
  );
  await verifyMraiNativeContentStreamCandidate(sourceBytes, prepared);
  const record = JSON.parse(JSON.stringify({
    id: input.transactionId,
    schemaVersion: MRAI_NATIVE_FONT_EXTENSION_TEXT_SCHEMA_VERSION,
    transactionSchemaVersion: prepared.schemaVersion,
    kind: "font-extension-text-replacement",
    featureId: PDF_NATIVE_FONT_EXTENSION_TEXT_FEATURE_ID,
    controlId: PDF_NATIVE_FONT_EXTENSION_TEXT_CONTROL_ID,
    action: MRAI_NATIVE_FONT_EXTENSION_TEXT_ACTION,
    transactionId: input.transactionId,
    baseSessionRevision: input.baseSessionRevision,
    selectionRevision: input.selection.selectionRevision,
    source: { resourceId: input.source.resourceId, resourceVersionId: input.source.resourceVersionId, sha256: sourceSha256 },
    target: {
      objectNumber: request.identity.objectNumber,
      generation: request.identity.generation,
      expectedObjectSha256: request.expectedObjectSha256,
      expectedStreamSha256: request.expectedStreamSha256,
    },
    operatorSpan: prepared.request.operatorSpan,
    replacement: [...request.replacement],
    replacementSha256: prepared.replacementSha256,
    ...pdfNativeStructActualTextRecordField(prepared.request),
    ...pdfNativePdfaIdentificationRecordField(prepared),
    untouchedObjects: prepared.request.untouchedObjects,
    formInstanceIsolation: null,
    pageExtGStateMutation: null,
    pageContentConsolidation: consolidation,
    pageFontExtensionMutation: {
      ...pageFontExtensionMutation,
      fontPackage: serializeFontMutationPackage(pageFontExtensionMutation.fontPackage),
      ...(pageFontExtensionMutation.additionalFontPackages
        ? { additionalFontPackages: pageFontExtensionMutation.additionalFontPackages.map(serializePdfNativeFontExtensionPackage) }
        : {}),
    },
    derivedCandidate: {
      sha256: prepared.candidateSha256,
      byteLength: prepared.candidateBytes.length,
      sourceLength: prepared.sourceLength,
      sourceXrefOffset: prepared.sourceXrefOffset,
      changedObjectOffset: prepared.changedObjectOffset,
      incrementalXrefOffset: prepared.incrementalXrefOffset,
      changedObjectSha256: prepared.changedObjectSha256,
      changedStreamSha256: prepared.changedStreamSha256,
      revisedObjects: derivedObjects(prepared),
    },
    listFidelity: null,
    textFidelity: null,
    embeddedImageRemoval: null,
    embeddedImageReplacement: null,
    outlineMutation: null,
    vectorGeometry: null,
    vectorDeletion: null,
    vectorStyle: null,
    fontExtensionText: {
      schemaVersion: MRAI_NATIVE_FONT_EXTENSION_TEXT_SCHEMA_VERSION,
      sourceFidelity: plan.evidence,
      replacementText: plan.evidence.replacementDisplayText,
    },
  })) as PdfSharedStateRecord;
  const sourceProof: PdfNativeListTextSelectionProof = Object.freeze({
    target: Object.freeze({ objectNumber: request.identity.objectNumber, generation: request.identity.generation }),
    evidence: plan.evidence,
    properties: null,
  });
  return Object.freeze({ sourceProof, prepared, record });
}

/** The library a write is planned with when the editor hands the slices
 * over (the engine host, which fetches nothing): each is kept only when its
 * bytes are the ones the lock pins. */
export function pdfNativeEmbeddedFontLibraryFromPrograms(
  programs: readonly Readonly<{ sha256: string; bytes: Uint8Array; pageFont?: PdfPageFont }>[],
  /** The served font list as the editor fetched it (narrowed to the
   * draft's families): a planner with no page to ask for it (the browser's
   * plan Worker) plans from this one; a host that holds the list loads it. */
  manifest?: EmbeddableFontManifest | null,
): Promise<PdfNativeEmbeddedFontLibrary> {
  const held = new Map(programs.map((program) => [program.sha256, program.bytes]));
  // A page font's program is checked against its own digest and size, as a
  // slice is (pdfPageFontSlice); the lock knows nothing of it.
  const pageFonts = programs.flatMap((program) => {
    const font = program.pageFont ? readPdfPageFont(program.pageFont) : null;
    return font && font.sha256 === program.sha256 ? [font] : [];
  });
  return (manifest ? Promise.resolve(manifest) : loadEmbeddableFontManifest()).then((manifest) => ({
    manifest,
    ...(pageFonts.length ? { pageFonts } : {}),
    fingerprint: JSON.stringify([[...held.keys()].sort(), manifest, pageFonts.map((font) => font.sha256).sort()]),
    loadSlice: async (slice: EmbeddableFontSlice) => {
      const bytes = held.get(slice.sha256);
      const verified = bytes ? await verifyEmbeddableFontSlice(slice, bytes) : null;
      if (!verified) throw new PdfNativeEmbeddedFontError("unavailable", "");
      return verified;
    },
  }));
}

/** planPdfNativeEmbeddedFontPropertyEdit as the engine host runs it for the
 * editor's preview: the served slices travel with the call. */
export async function planPdfNativeEmbeddedFontPropertyEditFromPrograms(
  input: Omit<Parameters<typeof planPdfNativeEmbeddedFontPropertyEdit>[0], "library"> & {
    readonly programs: readonly Readonly<{ sha256: string; bytes: Uint8Array; pageFont?: PdfPageFont }>[];
    /** The served font list the slices come from, narrowed to the draft's
     * families: a planner that cannot ask for it (the browser's planner
     * Worker) plans from this one; a host that holds the list loads it. */
    readonly manifest?: EmbeddableFontManifest | null;
  },
): Promise<PdfNativeFontExtensionTextPlan> {
  const { programs, manifest, ...rest } = input;
  return planPdfNativeEmbeddedFontPropertyEdit({
    ...rest,
    library: await pdfNativeEmbeddedFontLibraryFromPrograms(programs, manifest ?? null),
  });
}

export async function planPdfNativeFontReuseTextCandidate(input: {
  readonly sourceBytes: Uint8Array;
  readonly selection: PdfNativeListTextSelectionEvidence;
  readonly replacementText: string;
  readonly fontBytes: Uint8Array | ArrayBuffer;
  readonly reusable?: PdfNativeReusablePageFontExtension;
}): Promise<PdfNativeFontReuseTextCandidatePlan> {
  if (
    !input.replacementText ||
    input.replacementText.length > MAX_REPLACEMENT_CODE_UNITS ||
    /[\u0000-\u001f\u007f]/u.test(input.replacementText)
  ) {
    return invalid("Invalid native font reuse request.");
  }
  const sourceBytes = ownedPdfSourceBytes(input.sourceBytes);
  const sourceSha256 = await sha256PdfBytes(sourceBytes);
  if (sourceSha256 !== input.selection.sourceSha256.toLowerCase()) {
    return invalid("Native font reuse source is stale.");
  }
  const sourceProof = await proveMraiNativeListTextSelectionForWriter({
    sourceBytes,
    selection: input.selection,
  });
  const fidelity = sourceProof.evidence;
  const layout = fidelity.sameLineLayoutState;
  if (
    fidelity.schemaVersion !== MRAI_NATIVE_LIST_TEXT_SCHEMA_VERSION ||
    fidelity.wrapping !== "same-line-only" ||
    fidelity.linePlan !== null ||
    fidelity.selection?.items !== undefined ||
    // One show replaces the proof's whole span: a line of several text
    // objects would lose the state its later objects set (the run route
    // writes it run by run instead, through an alias of the font).
    fidelity.positionedRunGroup !== undefined ||
    fidelity.consolidatedRuns !== undefined ||
    fidelity.positionedLines !== undefined ||
    fidelity.lineShift !== undefined ||
    fidelity.originalText !== fidelity.replacementText ||
    fidelity.originalCodeBytes.join(",") !==
      fidelity.replacementCodeBytes.join(",") ||
    fidelity.originalText === input.replacementText ||
    !layout
  ) {
    return invalid("Native font reuse requires one exact source run.");
  }
  const reuse =
    input.reusable ??
    (await inspectMraiNativeReusablePageFontExtension({
      sourceBytes,
      pageNumber: input.selection.pageNumber,
      replacementText: input.replacementText,
      fontBytes: input.fontBytes,
    }));
  if (
    !reuse ||
    reuse.sourceSha256 !== sourceSha256 ||
    reuse.pageNumber !== input.selection.pageNumber ||
    reuse.currentToUnicode.objectSha256 !==
      reuse.replacementToUnicode.objectSha256
  ) {
    return invalid("Native font reuse requires a compatible ToUnicode map.");
  }
  const glyphCount = reuse.fontPackage.encodedHex.length / 4;
  const replacementAdvance =
    ((reuse.fontPackage.widthAt1000 * fidelity.fontSize) / 1_000 +
      layout.characterSpacing * glyphCount) *
    (layout.horizontalScale / 100);
  if (replacementAdvance > fidelity.originalAdvance + 0.001) {
    return invalid("Native font reuse replacement exceeds its text run.");
  }
  const page = await inspectMraiNativePageContentGraph(
    sourceBytes,
    input.selection.pageNumber,
  );
  const stream = await inspectMraiNativeContentStream(
    sourceBytes,
    sourceProof.target,
  );
  if (
    !page.contentStreams.some(
      (candidate) =>
        candidate.objectNumber === stream.identity.objectNumber &&
        candidate.generation === stream.identity.generation,
    )
  ) {
    return invalid("Native font reuse stream ownership is ambiguous.");
  }
  const size = formatNumber(fidelity.fontSize);
  const replacement = ascii(
    `/${reuse.resourceName} ${size} Tf <${reuse.fontPackage.encodedHex}> Tj /${fidelity.resourceName} ${size} Tf`,
  );
  const duplicateSpans = fontSetDuplicateSpans(sourceProof, fidelity.textOperatorSpan, true);
  const resourceOwner = await inspectMraiNativeObject(
    sourceBytes,
    page.resourceAuthority.owner,
  );
  const untouched = [
    {
      objectNumber: resourceOwner.identity.objectNumber,
      generation: resourceOwner.identity.generation,
      expectedObjectSha256: resourceOwner.objectSha256,
      role: "resource" as const,
    },
    ...reuse.sharedObjects.map((object) => ({
      objectNumber: object.objectNumber,
      generation: object.generation,
      expectedObjectSha256: object.sha256,
      role: "resource" as const,
    })),
    {
      objectNumber: reuse.currentToUnicode.objectNumber,
      generation: reuse.currentToUnicode.generation,
      expectedObjectSha256: reuse.currentToUnicode.objectSha256,
      role: "resource" as const,
    },
  ];
  const prepared = await prepareMraiNativeContentStreamTransaction(
    sourceBytes,
    // The reused face is the page's own: it reads by its own map.
    await withStructActualText(sourceBytes, page.pageNumber, {
      identity: stream.identity,
      expectedObjectSha256: stream.objectSha256,
      expectedStreamSha256: stream.streamSha256,
      operatorSpan: {
        start: fidelity.textOperatorSpan.start,
        end: fidelity.textOperatorSpan.end,
        expectedSha256: fidelity.textOperatorSpan.sha256,
      },
      replacement,
      // A line the page draws again in its place is set in every copy.
      ...(duplicateSpans ? { duplicateSpans } : {}),
      untouchedObjects: withProvenResourcePreconditions(untouched, sourceProof),
    }, [], { before: input.selection.item.str, after: input.replacementText }),
  );
  await verifyMraiNativeContentStreamCandidate(sourceBytes, prepared);
  return Object.freeze({ sourceProof, prepared, reuse });
}

export async function planPdfNativeFontReuseTextEdit(input: {
  readonly sourceBytes: Uint8Array;
  readonly source: PdfNativeFontExtensionSourceIdentity;
  readonly selection: PdfNativeListTextSelectionEvidence;
  readonly replacementText: string;
  readonly fontBytes: Uint8Array | ArrayBuffer;
  readonly transactionId: string;
  readonly baseSessionRevision: number;
  readonly reusable?: PdfNativeReusablePageFontExtension;
}): Promise<PdfNativeFontReuseTextPlan> {
  if (
    !input.source.resourceId ||
    !input.source.resourceVersionId ||
    !input.transactionId ||
    !Number.isSafeInteger(input.baseSessionRevision) ||
    input.baseSessionRevision < 0
  ) {
    return invalid("Invalid native font reuse request.");
  }
  const candidate = await planPdfNativeFontReuseTextCandidate(input);
  const prepared = candidate.prepared;
  const recordValue = JSON.parse(
    JSON.stringify({
      id: input.transactionId,
      schemaVersion: MRAI_NATIVE_FONT_REUSE_TEXT_SCHEMA_VERSION,
      transactionSchemaVersion: prepared.schemaVersion,
      kind: "font-reuse-text-replacement",
      featureId: PDF_NATIVE_FONT_EXTENSION_TEXT_FEATURE_ID,
      controlId: PDF_NATIVE_FONT_EXTENSION_TEXT_CONTROL_ID,
      action: MRAI_NATIVE_FONT_REUSE_TEXT_ACTION,
      transactionId: input.transactionId,
      baseSessionRevision: input.baseSessionRevision,
      selectionRevision: input.selection.selectionRevision,
      source: {
        resourceId: input.source.resourceId,
        resourceVersionId: input.source.resourceVersionId,
        sha256: candidate.reuse.sourceSha256,
      },
      target: {
        objectNumber: prepared.request.identity.objectNumber,
        generation: prepared.request.identity.generation,
        expectedObjectSha256: prepared.request.expectedObjectSha256,
        expectedStreamSha256: prepared.request.expectedStreamSha256,
      },
      operatorSpan: prepared.request.operatorSpan,
      replacement: [...prepared.request.replacement],
      // The copies of the show the page draws again in its place.
      ...(prepared.request.duplicateSpans ? { duplicateSpans: prepared.request.duplicateSpans } : {}),
      replacementSha256: prepared.replacementSha256,
      ...pdfNativeStructActualTextRecordField(prepared.request),
      ...pdfNativePdfaIdentificationRecordField(prepared),
      untouchedObjects: prepared.request.untouchedObjects,
      formInstanceIsolation: null,
      pageExtGStateMutation: null,
      pageFontExtensionMutation: null,
      derivedCandidate: {
        sha256: prepared.candidateSha256,
        byteLength: prepared.candidateBytes.length,
        sourceLength: prepared.sourceLength,
        sourceXrefOffset: prepared.sourceXrefOffset,
        changedObjectOffset: prepared.changedObjectOffset,
        incrementalXrefOffset: prepared.incrementalXrefOffset,
        changedObjectSha256: prepared.changedObjectSha256,
        changedStreamSha256: prepared.changedStreamSha256,
      },
      listFidelity: null,
      textFidelity: null,
      embeddedImageRemoval: null,
      embeddedImageReplacement: null,
      outlineMutation: null,
      vectorGeometry: null,
      vectorDeletion: null,
      vectorStyle: null,
      fontExtensionText: null,
      fontReuseText: {
        schemaVersion: MRAI_NATIVE_FONT_REUSE_TEXT_SCHEMA_VERSION,
        sourceFidelity: candidate.sourceProof.evidence,
        replacementText: input.replacementText,
        pageNumber: candidate.reuse.pageNumber,
        resourceName: candidate.reuse.resourceName,
        fontEncoding: {
          kind: "identity-h-cidfonttype2",
          codeUnitBytes: 2,
        },
        resources: candidate.reuse.resources,
      },
      fontExtensionHistory: null,
    }),
  ) as PdfSharedStateRecord;
  return Object.freeze({ ...candidate, record: recordValue });
}

function fontExtensionPackage(value: unknown): PdfNativeFontExtensionPackage {
  const wire = record(value);
  const objects =
    Array.isArray(wire.objects) &&
    wire.objects.length >= 4 &&
    wire.objects.length <= 16
      ? wire.objects.map((entry) => {
          const item = record(entry);
          return Object.freeze({
            objectNumber: positiveInteger(item.objectNumber, "font object"),
            generation: nonnegativeInteger(
              item.generation,
              "font generation",
            ) as 0,
            sha256: sha(item.sha256, "font object SHA"),
            bytes: base64Bytes(pdfNativeFontObjectWireBase64(item)),
          });
        })
      : invalid("Native font extension package objects are invalid.");
  const font = record(wire.font);
  return Object.freeze({
    schemaVersion:
      wire.schemaVersion === MRAI_NATIVE_FONT_EXTENSION_PACKAGE_SCHEMA_VERSION
        ? wire.schemaVersion
        : invalid("Native font extension package schema is invalid."),
    fontSha256: sha(
      wire.fontSha256,
      "font SHA",
    ) as PdfNativeFontExtensionPackage["fontSha256"],
    ...(wire.programFace === undefined
      ? {}
      : (() => {
          const face = record(wire.programFace);
          return {
            programFace: Object.freeze({
              family: typeof face.family === "string" && face.family
                ? face.family
                : invalid("Native font extension program face is invalid."),
              weight: positiveInteger(face.weight, "program face weight"),
              italic: typeof face.italic === "boolean"
                ? face.italic
                : invalid("Native font extension program face is invalid."),
              ...(face.page === undefined
                ? {}
                : face.page === true
                  ? { page: true as const }
                  : invalid("Native font extension program face is invalid.")),
            }),
          };
        })()),
    resourceName:
      typeof wire.resourceName === "string"
        ? wire.resourceName
        : invalid("Native font extension resource name is invalid."),
    font: Object.freeze({
      objectNumber: positiveInteger(font.objectNumber, "font reference"),
      generation: nonnegativeInteger(
        font.generation,
        "font reference generation",
      ) as 0,
    }),
    encodedHex:
      typeof wire.encodedHex === "string"
        ? wire.encodedHex
        : invalid("Native font extension encoded text is invalid."),
    widthAt1000:
      typeof wire.widthAt1000 === "number"
        ? wire.widthAt1000
        : invalid("Native font extension width is invalid."),
    unitsPerEm: positiveInteger(wire.unitsPerEm, "units per em"),
    ascent:
      typeof wire.ascent === "number"
        ? wire.ascent
        : invalid("Native font extension ascent is invalid."),
    descent:
      typeof wire.descent === "number"
        ? wire.descent
        : invalid("Native font extension descent is invalid."),
    objectByteLength: positiveInteger(
      wire.objectByteLength,
      "package byte length",
    ),
    objects: Object.freeze(objects),
  });
}

function fontAliasPackage(value: unknown): PdfNativeFontAliasPackage {
  const wire = record(value);
  const font = record(wire.font);
  const sourceFont = record(wire.sourceFont);
  const sourceToUnicode = record(wire.sourceToUnicode);
  const objects =
    Array.isArray(wire.objects) && wire.objects.length === 2
      ? wire.objects.map((entry) => {
          const item = record(entry);
          return Object.freeze({
            objectNumber: positiveInteger(item.objectNumber, "alias object"),
            generation: nonnegativeInteger(
              item.generation,
              "alias generation",
            ) as 0,
            sha256: sha(item.sha256, "alias object SHA"),
            bytes: base64Bytes(pdfNativeFontObjectWireBase64(item)),
          });
        })
      : invalid("Native font alias objects are invalid.");
  return Object.freeze({
    schemaVersion:
      wire.schemaVersion === "mrai.native-font-alias-package.v1"
        ? wire.schemaVersion
        : invalid("Native font alias schema is invalid."),
    fontSha256: sha(
      wire.fontSha256,
      "font alias SHA",
    ) as PdfNativeFontAliasPackage["fontSha256"],
    resourceName:
      typeof wire.resourceName === "string"
        ? wire.resourceName
        : invalid("Native font alias resource name is invalid."),
    font: Object.freeze({
      objectNumber: positiveInteger(font.objectNumber, "alias font"),
      generation: nonnegativeInteger(font.generation, "alias font generation") as 0,
    }),
    sourceFont: Object.freeze({
      objectNumber: positiveInteger(sourceFont.objectNumber, "source font"),
      generation: nonnegativeInteger(sourceFont.generation, "source font generation"),
      expectedObjectSha256: sha(sourceFont.expectedObjectSha256, "source font SHA"),
    }),
    sourceToUnicode: Object.freeze({
      objectNumber: positiveInteger(sourceToUnicode.objectNumber, "source CMap"),
      generation: nonnegativeInteger(
        sourceToUnicode.generation,
        "source CMap generation",
      ),
      expectedObjectSha256: sha(
        sourceToUnicode.expectedObjectSha256,
        "source CMap SHA",
      ),
    }),
    encodedHex:
      typeof wire.encodedHex === "string"
        ? wire.encodedHex
        : invalid("Native font alias text is invalid."),
    widthAt1000:
      typeof wire.widthAt1000 === "number"
        ? wire.widthAt1000
        : invalid("Native font alias width is invalid."),
    objectByteLength: positiveInteger(
      wire.objectByteLength,
      "alias byte length",
    ),
    objects: Object.freeze(objects),
  });
}

function fontMutationPackage(value: unknown) {
  return record(value).schemaVersion === "mrai.native-font-alias-package.v1"
    ? fontAliasPackage(value)
    : fontExtensionPackage(value);
}

function expectedObject(value: unknown) {
  const item = record(value);
  return Object.freeze({
    objectNumber: positiveInteger(item.objectNumber, "object number"),
    generation: nonnegativeInteger(item.generation, "object generation"),
    expectedObjectSha256: sha(item.expectedObjectSha256, "object SHA"),
  });
}

/** The structure elements a forward record wrote beside its text, as the
 * transaction request names them (the origin is prepared again with them). */
function originStructActualText(
  origin: Record<string, unknown>,
): { structActualText?: readonly PdfNativeStructActualTextRevision[] } {
  if (origin.structActualText === undefined || origin.structActualText === null) return {};
  if (!Array.isArray(origin.structActualText)) return invalid("Native font extension structure elements are invalid.");
  return {
    structActualText: Object.freeze(origin.structActualText.map((entry) => {
      const item = record(entry);
      const page = record(item.page);
      return Object.freeze({
        objectNumber: positiveInteger(item.objectNumber, "structure element"),
        generation: nonnegativeInteger(item.generation, "structure element generation"),
        expectedObjectSha256: sha(item.expectedObjectSha256, "structure element SHA"),
        page: Object.freeze({
          objectNumber: positiveInteger(page.objectNumber, "structure page"),
          generation: nonnegativeInteger(page.generation, "structure page generation"),
        }),
        mcids: Array.isArray(item.mcids)
          ? Object.freeze(item.mcids.map((mcid) => nonnegativeInteger(mcid, "MCID")))
          : invalid("Native font extension structure MCIDs are invalid."),
        previous: bytes(item.previous, "structure token"),
        value: bytes(item.value, "structure token"),
        text: typeof item.text === "string" ? item.text : invalid("Native font extension structure text is invalid."),
      });
    })),
  };
}

async function replayExtensionOriginPlan(input: {
  readonly currentSourceBytes: Uint8Array;
  readonly originRecord: PdfSharedStateRecord;
  readonly resourceId: string;
}): Promise<{
  originSourceBytes: Uint8Array;
  prepared: PdfNativePreparedContentStreamTransaction;
  historyKind: "extension";
  page: Readonly<{ objectNumber: number; generation: number }>;
}> {
  const origin = record(input.originRecord);
  const source = record(origin.source);
  const target = record(origin.target);
  const span = record(origin.operatorSpan);
  const mutation = record(origin.pageFontExtensionMutation);
  const category = record(mutation.category);
  const derived = record(origin.derivedCandidate);
  if (
    origin.schemaVersion !== MRAI_NATIVE_FONT_EXTENSION_TEXT_SCHEMA_VERSION ||
    origin.action !== MRAI_NATIVE_FONT_EXTENSION_TEXT_ACTION ||
    origin.kind !== "font-extension-text-replacement" ||
    origin.id !== origin.transactionId ||
    source.resourceId !== input.resourceId ||
    // A font installed in an isolated Form's own resources is not a page
    // font: nothing on the page can reuse it.
    origin.formInstanceIsolation !== null ||
    mutation.form !== undefined
  ) {
    return invalid("Native font extension history origin is invalid.");
  }
  const originSourceLength = positiveInteger(
    derived.sourceLength,
    "origin source length",
  );
  const originCandidateLength = positiveInteger(
    derived.byteLength,
    "origin candidate length",
  );
  if (
    originCandidateLength <= originSourceLength ||
    input.currentSourceBytes.length < originCandidateLength
  ) {
    return invalid("Native font extension history prefix is invalid.");
  }
  const originSourceBytes = input.currentSourceBytes.slice(
    0,
    originSourceLength,
  );
  const originCandidateBytes = input.currentSourceBytes.slice(
    0,
    originCandidateLength,
  );
  if (
    (await sha256PdfBytes(originSourceBytes)) !==
      sha(source.sha256, "origin source SHA") ||
    (await sha256PdfBytes(originCandidateBytes)) !==
      sha(derived.sha256, "origin candidate SHA")
  ) {
    return invalid("Native font extension history prefix changed.");
  }
  const categoryPlan: PdfNativePageFontExtensionMutationPlan["category"] =
    category.kind === "indirect"
      ? Object.freeze({
          kind: "indirect" as const,
          expectedDictionarySha256: sha(
            category.expectedDictionarySha256,
            "Font category SHA",
          ),
          target: expectedObject(category.target),
        })
      : category.kind === "missing" || category.kind === "direct"
        ? Object.freeze({
            kind: category.kind,
            expectedDictionarySha256: sha(
              category.expectedDictionarySha256,
              "Font category SHA",
            ),
          })
        : invalid("Native font extension Font category is invalid.");
  const pageFontExtensionMutation: PdfNativePageFontExtensionMutationPlan =
    Object.freeze({
      pageNumber: positiveInteger(mutation.pageNumber, "page number"),
      page: expectedObject(mutation.page),
      resourceOwner: Object.freeze({
        ...expectedObject(mutation.resourceOwner),
        dictionaryKind:
          record(mutation.resourceOwner).dictionaryKind ===
            "nested-resources" ||
          record(mutation.resourceOwner).dictionaryKind ===
            "resource-dictionary"
            ? (record(mutation.resourceOwner).dictionaryKind as
                "nested-resources" | "resource-dictionary")
            : invalid("Native font extension resource owner is invalid."),
      }),
      category: categoryPlan,
      fontPackage: fontMutationPackage(mutation.fontPackage),
      ...(mutation.additionalFontPackages === undefined || mutation.additionalFontPackages === null
        ? {}
        : {
            additionalFontPackages: Array.isArray(mutation.additionalFontPackages)
              ? Object.freeze(mutation.additionalFontPackages.map(fontExtensionPackage))
              : invalid("Native font extension additional packages are invalid."),
          }),
    });
  const prepared = await prepareMraiNativeContentStreamTransaction(
    originSourceBytes,
    {
      identity: {
        sourceSha256: sha(source.sha256, "origin source SHA"),
        objectNumber: positiveInteger(target.objectNumber, "target object"),
        generation: nonnegativeInteger(target.generation, "target generation"),
      },
      expectedObjectSha256: sha(
        target.expectedObjectSha256,
        "target object SHA",
      ),
      expectedStreamSha256: sha(
        target.expectedStreamSha256,
        "target stream SHA",
      ),
      operatorSpan: {
        start: nonnegativeInteger(span.start, "operator start"),
        end: positiveInteger(span.end, "operator end"),
        expectedSha256: sha(span.expectedSha256, "operator SHA"),
      },
      replacement: bytes(origin.replacement, "replacement bytes"),
      ...originDuplicateSpans(origin),
      untouchedObjects: Array.isArray(origin.untouchedObjects)
        ? origin.untouchedObjects.map((entry) => {
            const untouched = expectedObject(entry);
            const role = record(entry).role;
            return Object.freeze({
              ...untouched,
              role:
                role === "resource" || role === "untouched-object"
                  ? role
                  : invalid("Native font extension untouched role is invalid."),
            });
          })
        : invalid("Native font extension untouched objects are invalid."),
      pageFontExtensionMutation,
      // The rules a property edit wrote beside its run, written again.
      ...(Array.isArray(origin.textRules) && origin.textRules.length
        ? {
            textRules: origin.textRules as unknown as NonNullable<
              Parameters<typeof prepareMraiNativeContentStreamTransaction>[1]["textRules"]
            >,
          }
        : {}),
      // A page drawn from several streams was edited through the one it
      // consolidates them into; the origin is replayed the same way.
      ...(origin.pageContentConsolidation
        ? {
            pageContentConsolidation:
              origin.pageContentConsolidation as unknown as NonNullable<
                Parameters<typeof prepareMraiNativeContentStreamTransaction>[1]["pageContentConsolidation"]
              >,
          }
        : {}),
      ...originStructActualText(origin),
    },
  );
  await verifyMraiNativeContentStreamCandidate(originSourceBytes, prepared);
  if (
    prepared.candidateSha256 !== sha(derived.sha256, "origin candidate SHA") ||
    prepared.candidateBytes.length !== originCandidateLength ||
    prepared.sourceXrefOffset !==
      nonnegativeInteger(derived.sourceXrefOffset, "origin xref") ||
    prepared.incrementalXrefOffset !==
      nonnegativeInteger(derived.incrementalXrefOffset, "candidate xref") ||
    !prepared.candidateBytes.every(
      (value, index) => value === originCandidateBytes[index],
    )
  ) {
    return invalid("Native font extension origin replay changed.");
  }
  return Object.freeze({
    originSourceBytes,
    prepared,
    historyKind: "extension" as const,
    page: Object.freeze({
      objectNumber: pageFontExtensionMutation.page.objectNumber,
      generation: pageFontExtensionMutation.page.generation,
    }),
  });
}

async function replayFontReuseOriginPlan(input: {
  readonly currentSourceBytes: Uint8Array;
  readonly originRecord: PdfSharedStateRecord;
  readonly resourceId: string;
}): Promise<{
  originSourceBytes: Uint8Array;
  prepared: PdfNativePreparedContentStreamTransaction;
  historyKind: "reuse";
  page: Readonly<{ objectNumber: number; generation: number }>;
}> {
  const origin = record(input.originRecord);
  const source = record(origin.source);
  const target = record(origin.target);
  const span = record(origin.operatorSpan);
  const evidence = record(origin.fontReuseText);
  const resources = record(evidence.resources);
  const page = record(resources.page);
  const derived = record(origin.derivedCandidate);
  if (
    origin.schemaVersion !== MRAI_NATIVE_FONT_REUSE_TEXT_SCHEMA_VERSION ||
    origin.action !== MRAI_NATIVE_FONT_REUSE_TEXT_ACTION ||
    origin.kind !== "font-reuse-text-replacement" ||
    origin.id !== origin.transactionId ||
    source.resourceId !== input.resourceId ||
    origin.pageFontExtensionMutation !== null ||
    evidence.schemaVersion !== MRAI_NATIVE_FONT_REUSE_TEXT_SCHEMA_VERSION
  ) {
    return invalid("Native font reuse history origin is invalid.");
  }
  const originSourceLength = positiveInteger(
    derived.sourceLength,
    "reuse origin source length",
  );
  const originCandidateLength = positiveInteger(
    derived.byteLength,
    "reuse origin candidate length",
  );
  if (
    originCandidateLength <= originSourceLength ||
    input.currentSourceBytes.length < originCandidateLength
  ) {
    return invalid("Native font reuse history prefix is invalid.");
  }
  const originSourceBytes = input.currentSourceBytes.slice(
    0,
    originSourceLength,
  );
  const originCandidateBytes = input.currentSourceBytes.slice(
    0,
    originCandidateLength,
  );
  if (
    (await sha256PdfBytes(originSourceBytes)) !==
      sha(source.sha256, "reuse origin source SHA") ||
    (await sha256PdfBytes(originCandidateBytes)) !==
      sha(derived.sha256, "reuse origin candidate SHA")
  ) {
    return invalid("Native font reuse history prefix changed.");
  }
  const prepared = await prepareMraiNativeContentStreamTransaction(
    originSourceBytes,
    {
      identity: {
        sourceSha256: sha(source.sha256, "reuse origin source SHA"),
        objectNumber: positiveInteger(target.objectNumber, "reuse target"),
        generation: nonnegativeInteger(
          target.generation,
          "reuse target generation",
        ),
      },
      expectedObjectSha256: sha(
        target.expectedObjectSha256,
        "reuse target object SHA",
      ),
      expectedStreamSha256: sha(
        target.expectedStreamSha256,
        "reuse target stream SHA",
      ),
      operatorSpan: {
        start: nonnegativeInteger(span.start, "reuse operator start"),
        end: positiveInteger(span.end, "reuse operator end"),
        expectedSha256: sha(span.expectedSha256, "reuse operator SHA"),
      },
      replacement: bytes(origin.replacement, "reuse replacement bytes"),
      ...originDuplicateSpans(origin),
      untouchedObjects: Array.isArray(origin.untouchedObjects)
        ? origin.untouchedObjects.map((entry) => {
            const untouched = expectedObject(entry);
            const role = record(entry).role;
            return Object.freeze({
              ...untouched,
              role:
                role === "resource" || role === "untouched-object"
                  ? role
                  : invalid("Native font reuse untouched role is invalid."),
            });
          })
        : invalid("Native font reuse untouched objects are invalid."),
      ...originStructActualText(origin),
    },
  );
  await verifyMraiNativeContentStreamCandidate(originSourceBytes, prepared);
  if (
    prepared.candidateSha256 !==
      sha(derived.sha256, "reuse origin candidate SHA") ||
    prepared.candidateBytes.length !== originCandidateLength ||
    !prepared.candidateBytes.every(
      (value, index) => value === originCandidateBytes[index],
    )
  ) {
    return invalid("Native font reuse origin replay changed.");
  }
  return Object.freeze({
    originSourceBytes,
    prepared,
    historyKind: "reuse" as const,
    page: Object.freeze({
      objectNumber: positiveInteger(page.objectNumber, "reuse page"),
      generation: nonnegativeInteger(page.generation, "reuse page generation"),
    }),
  });
}

async function replayOriginPlan(input: {
  readonly currentSourceBytes: Uint8Array;
  readonly originRecord: PdfSharedStateRecord;
  readonly resourceId: string;
}) {
  return input.originRecord.schemaVersion ===
    MRAI_NATIVE_FONT_REUSE_TEXT_SCHEMA_VERSION
    ? replayFontReuseOriginPlan(input)
    : replayExtensionOriginPlan(input);
}

function historyState(state: PdfNativeFontExtensionHistoryObjectState) {
  return Object.freeze({ ...state });
}

export async function planPdfNativeFontExtensionTextHistory(input: {
  readonly sourceBytes: Uint8Array;
  readonly source: PdfNativeFontExtensionSourceIdentity & {
    readonly sourceSha256: string;
  };
  readonly originRecord: PdfSharedStateRecord;
  readonly direction: "undo" | "redo";
  readonly transactionId: string;
  readonly baseSessionRevision: number;
}): Promise<PdfNativeFontExtensionTextHistoryPlan> {
  if (
    !input.source.resourceId ||
    !input.source.resourceVersionId ||
    !input.transactionId ||
    !Number.isSafeInteger(input.baseSessionRevision) ||
    input.baseSessionRevision < 0 ||
    (input.direction !== "undo" && input.direction !== "redo")
  ) {
    return invalid("Invalid native font extension history request.");
  }
  // A private copy: the history plan scrubs its bytes when it is done.
  const sourceBytes = Uint8Array.from(input.sourceBytes);
  const sourceSha256 = await sha256PdfBytes(sourceBytes);
  if (sourceSha256 !== sha(input.source.sourceSha256, "history source SHA")) {
    sourceBytes.fill(0);
    return invalid("Native font extension history source is stale.");
  }
  let prepared: PdfNativePreparedFontExtensionHistory;
  let lineage: Readonly<{
    transactionSchemaVersion: string;
    originSourceSha256: string;
    originCandidateSha256: string;
    originKind: "extension" | "reuse";
  }>;
  const kept = isPdfNativeObjectRestoreEvidence(input.originRecord)
    ? input.originRecord
    : null;
  if (kept) {
    // An origin a source consolidation compacted away: the objects its
    // history swaps are read from the evidence the consolidation record
    // kept for it (planMraiNativeFontHistoryFromEvidence).
    if (kept.resourceId !== input.source.resourceId) {
      sourceBytes.fill(0);
      return invalid("Native font extension history origin is invalid.");
    }
    try {
      prepared = await planMraiNativeFontHistoryFromEvidence({
        sourceBytes,
        kept: {
          originKind: kept.originKind,
          originSourceSha256: kept.originSourceSha256,
          target: kept.target,
          page: kept.page,
          originalTarget: restoreBytes(kept.originalTarget),
          extendedTarget: restoreBytes(kept.extendedTarget),
          originalPage: restoreBytes(kept.originalPage),
          extendedPage: restoreBytes(kept.extendedPage),
          originalTargetStreamSha256: kept.originalTargetStreamSha256,
          extendedTargetStreamSha256: kept.extendedTargetStreamSha256,
          protected: kept.protected,
        },
        direction: input.direction,
      });
    } catch (error) {
      sourceBytes.fill(0);
      throw error;
    }
    lineage = {
      transactionSchemaVersion: kept.transactionSchemaVersion,
      originSourceSha256: kept.originSourceSha256,
      originCandidateSha256: kept.originCandidateSha256,
      originKind: kept.originKind,
    };
  } else {
    const origin = await replayOriginPlan({
      currentSourceBytes: sourceBytes,
      originRecord: input.originRecord,
      resourceId: input.source.resourceId,
    });
    try {
      prepared =
        origin.historyKind === "reuse"
          ? await planMraiNativeFontReuseHistory({
              sourceBytes,
              originSourceBytes: origin.originSourceBytes,
              originPrepared: origin.prepared,
              page: origin.page,
              direction: input.direction,
            })
          : await planMraiNativePageFontExtensionHistory({
              sourceBytes,
              originSourceBytes: origin.originSourceBytes,
              originPrepared: origin.prepared,
              direction: input.direction,
            });
    } catch (error) {
      origin.prepared.candidateBytes.fill(0);
      origin.originSourceBytes.fill(0);
      sourceBytes.fill(0);
      throw error;
    }
    lineage = {
      transactionSchemaVersion: origin.prepared.schemaVersion,
      originSourceSha256: origin.prepared.request.identity.sourceSha256,
      originCandidateSha256: origin.prepared.candidateSha256,
      originKind: origin.historyKind,
    };
    origin.prepared.candidateBytes.fill(0);
    origin.originSourceBytes.fill(0);
  }
  // The structure elements the origin wrote beside its text go back with
  
  let structActualText: Awaited<ReturnType<typeof planMraiNativeStructActualTextHistory>>;
  try {
    structActualText = await planMraiNativeStructActualTextHistory(
      sourceBytes,
      (kept ? (kept as { structActualText?: unknown }).structActualText : record(input.originRecord).structActualText),
      input.direction,
    );
    if (structActualText) {
      const appended = await appendMraiNativeStructActualTextRevisions(sourceBytes, prepared, structActualText);
      prepared.candidateBytes.fill(0);
      prepared = Object.freeze({
        ...prepared,
        candidateBytes: appended.candidateBytes,
        candidateSha256: await sha256PdfBytes(appended.candidateBytes),
        incrementalXrefOffset: appended.incrementalXrefOffset,
        revisedObjects: Object.freeze([...prepared.revisedObjects, ...appended.objects].sort(
          (left, right) => left.objectNumber - right.objectNumber || left.generation - right.generation,
        )),
      });
    }
  } catch (error) {
    sourceBytes.fill(0);
    throw error;
  }
  const originId = String(input.originRecord.id);
  const recordValue = JSON.parse(
    JSON.stringify({
      id: input.transactionId,
      schemaVersion: MRAI_NATIVE_FONT_EXTENSION_TEXT_HISTORY_SCHEMA_VERSION,
      transactionSchemaVersion: lineage.transactionSchemaVersion,
      kind: "font-extension-text-history",
      featureId: PDF_NATIVE_FONT_EXTENSION_TEXT_FEATURE_ID,
      controlId: PDF_NATIVE_FONT_EXTENSION_TEXT_CONTROL_ID,
      action: MRAI_NATIVE_FONT_EXTENSION_TEXT_HISTORY_ACTION,
      transactionId: input.transactionId,
      baseSessionRevision: input.baseSessionRevision,
      selectionRevision: 0,
      source: {
        resourceId: input.source.resourceId,
        resourceVersionId: input.source.resourceVersionId,
        sha256: sourceSha256,
      },
      target: {
        objectNumber: prepared.target.objectNumber,
        generation: prepared.target.generation,
        expectedObjectSha256: prepared.current.targetObjectSha256,
        expectedStreamSha256: prepared.current.targetStreamSha256,
      },
      operatorSpan: null,
      replacement: null,
      replacementSha256: null,
      ...pdfNativeStructActualTextRecordField({ structActualText: structActualText ?? undefined }),
      untouchedObjects: [],
      formInstanceIsolation: null,
      pageExtGStateMutation: null,
      pageFontExtensionMutation: null,
      derivedCandidate: {
        sha256: prepared.candidateSha256,
        byteLength: prepared.candidateBytes.length,
        sourceLength: prepared.sourceLength,
        sourceXrefOffset: prepared.sourceXrefOffset,
        changedObjectOffset: prepared.changedObjectOffset,
        incrementalXrefOffset: prepared.incrementalXrefOffset,
        changedObjectSha256: prepared.changedObjectSha256,
        changedStreamSha256: prepared.changedStreamSha256,
        revisedObjects: prepared.revisedObjects.map((item) => ({ ...item })),
      },
      listFidelity: null,
      textFidelity: null,
      embeddedImageRemoval: null,
      embeddedImageReplacement: null,
      outlineMutation: null,
      vectorGeometry: null,
      vectorDeletion: null,
      vectorStyle: null,
      fontExtensionText: null,
      fontExtensionHistory: {
        schemaVersion: MRAI_NATIVE_FONT_EXTENSION_TEXT_HISTORY_SCHEMA_VERSION,
        originTransactionId: originId,
        originSourceSha256: lineage.originSourceSha256,
        originCandidateSha256: lineage.originCandidateSha256,
        originKind: lineage.originKind,
        direction: input.direction,
        page: { ...prepared.page },
        current: historyState(prepared.current),
        desired: historyState(prepared.desired),
      },
    }),
  ) as PdfSharedStateRecord;
  sourceBytes.fill(0);
  return Object.freeze({ prepared, record: recordValue });
}
