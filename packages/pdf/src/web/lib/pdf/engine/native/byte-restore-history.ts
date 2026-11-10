/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { LATIN1_EXACT } from "./latin1";
import { unzlibSync } from "fflate";
import {
  NOT_RESTORABLE,
  int,
  isPdfNativeRestoreEvidence,
  mraiNativeRecordIsByteRestorable,
  record,
  type PdfNativeRestoreEvidence,
  type PdfNativeRestoreWindow,
} from "./byte-restore-model";
export { mraiNativeRecordIsByteRestorable };
import {
  decodeMraiNativePdfStringToken,
  decodeMraiNativePdfTextString,
  inspectMraiNativeContentStream,
  inspectMraiNativeObject,
  PdfNativeTransactionError,
  readMraiNativeStructElementActualText,
  sha256PdfBytes,
  type PdfNativeContentStreamTransactionRequest,
  type PdfNativeStructActualTextRevision,
} from "./content-stream-transaction";
import type { PdfNativeTextWindowHistoryEvidence } from "./list-text-fidelity";
import { MRAI_NATIVE_BOX_MARK_CLOSE, MRAI_NATIVE_BOX_MARK_OPEN } from "./box-mark";

/** The history record's evidence schema (list-text-fidelity's
 * MRAI_NATIVE_TEXT_WINDOW_HISTORY_SCHEMA_VERSION; the text-window history was
 * the first byte restore and keeps its name on the wire). */
const HISTORY_SCHEMA_VERSION = "mrai.pdf-native-text.window-history.v1" as const;


function refuse(message: string): never {
  throw new PdfNativeTransactionError("unsupported-document", message);
}

export type Json = Record<string, unknown>;


function byteArray(value: unknown): Uint8Array | null {
  if (!Array.isArray(value) || value.some((entry) => !Number.isInteger(entry) || entry < 0 || entry > 255)) return null;
  return Uint8Array.from(value as number[]);
}


function concat(parts: readonly Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((total, part) => total + part.length, 0));
  let offset = 0;
  for (const part of parts) { out.set(part, offset); offset += part.length; }
  return out;
}

function equalBytes(left: Uint8Array, right: Uint8Array): boolean {
  if (left.length !== right.length) return false;
  for (let index = 0; index < left.length; index += 1) if (left[index] !== right[index]) return false;
  return true;
}

/** The replacement a record writes: its bytes, or its pieces put together
 * from the span they name (mirror of the server's `_record_replacement`). */
function recordReplacement(origin: Json, original: Uint8Array): Uint8Array {
  const pieces = origin.replacementPieces;
  if (!Array.isArray(pieces)) {
    const bytes = byteArray(origin.replacement);
    if (!bytes) refuse("The history's origin carries no replacement.");
    return bytes;
  }
  return concat(pieces.map((raw) => {
    const piece = record(raw);
    const literal = byteArray(piece?.literal);
    if (literal) return literal;
    const start = int(piece?.spanStart), end = int(piece?.spanEnd);
    if (start === null || end === null || end < start || end > original.length) refuse("The history's origin pieces are inconsistent.");
    return original.slice(start, end);
  }));
}

/** The window of the forward's source stream that holds every byte it
 * replaced, the same window in the stream it wrote, and that written stream
 * (mirror of the server's forward_edit_envelope). */
export async function forwardEditEnvelope(origin: Json, sourceStream: Uint8Array): Promise<{
  sourceWindow: readonly [number, number];
  writtenWindow: readonly [number, number];
  writtenStream: Uint8Array;
}> {
  const span = record(origin.operatorSpan);
  const start = int(span?.start), end = int(span?.end);
  if (start === null || end === null || end <= start || end > sourceStream.length) refuse("The history's origin span is outside its stream.");
  const original = sourceStream.slice(start, end);
  if (await sha256PdfBytes(original) !== span!.expectedSha256) refuse("The history's origin source bytes changed.");
  const written = recordReplacement(origin, original);
  if (await sha256PdfBytes(written) !== origin.replacementSha256) refuse("The history's origin replacement is inconsistent.");
  // A box mark's halves stand round the edit (mirror of replay). They are
  // part of the stream the forward wrote, not of its window: an undo writes
  // the edit's own bytes back inside the box's mark, which stays (the box
  // stays the one box it was), and a later edit of the box's position does
  // not move the window.
  const mark = record(origin.boxMark);
  const markStart = int(mark?.start), markEnd = int(mark?.end);
  if (mark && (markStart === null || markEnd === null || markStart > start || markEnd < end)) {
    refuse("The history's origin box mark is inconsistent.");
  }
  const edits: [number, number, Uint8Array][] = [
    ...(mark ? [[markStart!, markStart!, MRAI_NATIVE_BOX_MARK_OPEN] as [number, number, Uint8Array]] : []),
    [start, end, written],
  ];
  for (const raw of Array.isArray(origin.duplicateSpans) ? origin.duplicateSpans : []) {
    const copy = record(raw);
    const copyStart = int(copy?.start), copyEnd = int(copy?.end);
    if (copyStart === null || copyEnd === null) refuse("The history's origin copies are inconsistent.");
    const partStart = int(copy?.replacementStart), partEnd = int(copy?.replacementEnd);
    if (partStart !== null && partEnd !== null) {
      if (partEnd < partStart || partEnd > written.length) refuse("The history's origin copies are inconsistent.");
      edits.push([copyStart, copyEnd, written.slice(partStart, partEnd)]);
    } else edits.push([copyStart, copyEnd, written]);
  }
  // The rules a property edit rewrote beside the run go with it (mirror of
  // replay: each is one more span of the same stream).
  for (const raw of Array.isArray(origin.textRules) ? origin.textRules : []) {
    const rule = record(raw);
    const ruleStart = int(rule?.start), ruleEnd = int(rule?.end);
    if (ruleStart === null || ruleEnd === null || !Array.isArray(rule?.added)) refuse("The history's origin rules are inconsistent.");
    edits.push([ruleStart, ruleEnd, Uint8Array.from(rule!.added as number[])]);
  }
  if (mark) edits.push([markEnd!, markEnd!, MRAI_NATIVE_BOX_MARK_CLOSE]);
  edits.sort((left, right) => left[0] - right[0]);
  const parts: Uint8Array[] = [];
  let cursor = 0, delta = 0;
  for (const [editStart, editEnd, bytes] of edits) {
    if (editStart < cursor || editEnd > sourceStream.length || editEnd < editStart) refuse("The history's origin edits overlap.");
    parts.push(sourceStream.subarray(cursor, editStart), bytes);
    cursor = editEnd;
    delta += bytes.length - (editEnd - editStart);
  }
  parts.push(sourceStream.subarray(cursor));
  const own = mark ? edits.filter((edit) => edit[2] !== MRAI_NATIVE_BOX_MARK_OPEN && edit[2] !== MRAI_NATIVE_BOX_MARK_CLOSE) : edits;
  const first = own[0][0];
  const last = Math.max(...own.map((edit) => edit[1]));
  // The opening half stands before the window, the closing half after it.
  const shift = mark ? MRAI_NATIVE_BOX_MARK_OPEN.length : 0;
  const markDelta = mark ? MRAI_NATIVE_BOX_MARK_OPEN.length + MRAI_NATIVE_BOX_MARK_CLOSE.length : 0;
  return {
    sourceWindow: [first, last],
    writtenWindow: [first + shift, last + delta - markDelta + shift],
    writtenStream: concat(parts),
  };
}

/** Where `reference[window]` stands in `current`: in place when `current`
 * differs from `reference` only before the window or only after it, else the
 * one place in `current` that holds those bytes (mirror of the server's
 * current_window). A single byte names no place -- the `n` a vector delete
 * leaves is the letter of every other `n`, `scn` and `/Span` -- so a one-byte
 * window is found only in place or shifted, never searched for. */
export function currentWindow(reference: Uint8Array, window: readonly [number, number], current: Uint8Array): readonly [number, number] | null {
  const limit = Math.min(reference.length, current.length);
  let prefix = 0;
  while (prefix < limit && reference[prefix] === current[prefix]) prefix += 1;
  if (window[1] <= prefix) return window;
  let suffix = 0;
  while (suffix < limit - prefix && reference[reference.length - 1 - suffix] === current[current.length - 1 - suffix]) suffix += 1;
  if (window[0] >= reference.length - suffix) {
    const shift = current.length - reference.length;
    return [window[0] + shift, window[1] + shift];
  }
  const wanted = reference.subarray(window[0], window[1]);
  if (wanted.length < 2) return null;
  const first = indexOfBytes(current, wanted, 0);
  if (first < 0 || indexOfBytes(current, wanted, first + 1) >= 0) return null;
  return [first, first + wanted.length];
}

function indexOfBytes(haystack: Uint8Array, needle: Uint8Array, from: number): number {
  if (needle.length === 0) return -1;
  outer: for (let index = from; index + needle.length <= haystack.length; index += 1) {
    if (haystack[index] !== needle[0]) continue;
    for (let offset = 1; offset < needle.length; offset += 1) if (haystack[index + offset] !== needle[offset]) continue outer;
    return index;
  }
  return -1;
}

async function joinedSourceStream(before: Uint8Array, origin: Json): Promise<Uint8Array> {
  const consolidation = record(origin.pageContentConsolidation);
  const streams = Array.isArray(consolidation?.sourceStreams) ? consolidation!.sourceStreams : null;
  if (!consolidation || !streams?.length) refuse("The history's origin page streams are missing.");
  const parts: Uint8Array[] = [];
  for (const raw of streams) {
    const item = record(raw);
    const objectNumber = int(item?.objectNumber), generation = int(item?.generation);
    if (objectNumber === null || generation === null) refuse("The history's origin page streams are missing.");
    const content = await inspectMraiNativeContentStream(before, { objectNumber, generation });
    if (content.objectSha256 !== item!.expectedObjectSha256) refuse("The history's origin page stream changed.");
    parts.push(content.streamBytes);
  }
  const joined = concat(parts);
  if (await sha256PdfBytes(joined) !== consolidation.joinedDecodedSha256) refuse("The history's origin page streams changed.");
  return joined;
}

/** The objects a revision wrote, read from its own appended bytes. */
function objectsWrittenAfter(after: Uint8Array, sourceLength: number): { objectNumber: number; generation: number }[] {
  const text = LATIN1_EXACT.decode(after.subarray(sourceLength));
  const seen = new Set<string>();
  const out: { objectNumber: number; generation: number }[] = [];
  for (const match of text.matchAll(/(?:^|[\r\n])\s*(\d{1,10})\s+(\d{1,5})\s+obj\b/g)) {
    const key = `${match[1]} ${match[2]}`;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push({ objectNumber: Number(match[1]), generation: Number(match[2]) });
  }
  return out;
}

export interface PdfNativeByteRestoreHistoryInput {
  readonly sourceBytes: Uint8Array;
  readonly expectedSourceSha256: string;
  readonly resourceId: string;
  /** The forward record, as the journal holds it, or the restore evidence a
   * consolidation record kept for it. */
  readonly originRecord: Record<string, unknown>;
  readonly direction: "undo" | "redo";
}

export interface PdfNativeByteRestoreHistoryPlan {
  readonly request: PdfNativeContentStreamTransactionRequest;
  readonly history: PdfNativeTextWindowHistoryEvidence;
}









export async function planMraiNativeStructActualTextHistory(
  sourceBytes: Uint8Array,
  originStruct: unknown,
  direction: "undo" | "redo",
): Promise<readonly PdfNativeStructActualTextRevision[] | null> {
  if (originStruct === undefined || originStruct === null) return null;
  if (!Array.isArray(originStruct) || originStruct.length === 0) refuse("The history's origin structure elements are invalid.");
  const revisions: PdfNativeStructActualTextRevision[] = [];
  for (const raw of originStruct) {
    const item = record(raw), page = record(item?.page);
    const objectNumber = int(item?.objectNumber), generation = int(item?.generation);
    const pageNumber = int(page?.objectNumber), pageGeneration = int(page?.generation);
    const previous = byteArray(item?.previous), value = byteArray(item?.value);
    const mcids = Array.isArray(item?.mcids) ? (item!.mcids as unknown[]).map(int) : null;
    if (objectNumber === null || generation === null || pageNumber === null || pageGeneration === null ||
      !previous || !value || !mcids || mcids.some((mcid) => mcid === null)) {
      refuse("The history's origin structure elements are invalid.");
    }
    const [current, desired] = direction === "undo" ? [value, previous] : [previous, value];
    const element = await inspectMraiNativeObject(sourceBytes, { objectNumber, generation });
    const found = readMraiNativeStructElementActualText(element.objectBytes);
    if (!found || !equalBytes(found.token, current)) {
      refuse("The structure element no longer holds the text the edit wrote; undo the later edits first.");
    }
    const bytes = decodeMraiNativePdfStringToken(desired);
    const text = bytes ? decodeMraiNativePdfTextString(bytes, false) : null;
    if (text === null) refuse("The history's origin structure text is unreadable.");
    revisions.push(Object.freeze({
      objectNumber, generation,
      expectedObjectSha256: element.objectSha256,
      page: Object.freeze({ objectNumber: pageNumber, generation: pageGeneration }),
      mcids: Object.freeze(mcids as number[]),
      previous: current,
      value: desired,
      text,
    }));
  }
  return Object.freeze(revisions);
}

/** The bytes a restore evidence field names (base64 of a zlib stream). */
export function restoreBytes(value: string): Uint8Array {
  let packed: Uint8Array;
  try {
    const binary = atob(value);
    packed = Uint8Array.from(binary, (character) => character.charCodeAt(0));
  } catch {
    refuse("The history's kept bytes are unreadable.");
  }
  try {
    return unzlibSync(packed!);
  } catch {
    refuse("The history's kept bytes are unreadable.");
  }
}

/** Where the bytes a restore window held stand in `current`, read from the
 * window's digests instead of the stream they were found in (mirror of the
 * server's portable_window): in place when everything up to the window's end
 * is as it was, shifted when everything from its start is, else the one
 * place in `current` that holds them. A single byte is never searched for. */
export async function portableWindow(
  spec: PdfNativeRestoreWindow,
  wanted: Uint8Array,
  current: Uint8Array,
): Promise<readonly [number, number] | null> {
  const { start, end, length } = spec;
  if (end - start !== wanted.length) return null;
  if (current.length >= end && await sha256PdfBytes(current.subarray(0, end)) === spec.prefixSha256) return [start, end];
  const tail = length - start;
  if (current.length >= tail && await sha256PdfBytes(current.subarray(current.length - tail)) === spec.suffixSha256) {
    const shift = current.length - length;
    return [start + shift, end + shift];
  }
  if (wanted.length < 2) return null;
  const first = indexOfBytes(current, wanted, 0);
  if (first < 0 || indexOfBytes(current, wanted, first + 1) >= 0) return null;
  return [first, first + wanted.length];
}

/** The history of an origin a source consolidation compacted away: the
 * object its evidence names, the place its bytes stand now by the window's
 * digests, and the bytes it replaced (undo) or wrote (redo). */
async function planPortableByteRestore(
  input: PdfNativeByteRestoreHistoryInput,
  evidence: PdfNativeRestoreEvidence,
): Promise<PdfNativeByteRestoreHistoryPlan> {
  if (await sha256PdfBytes(input.sourceBytes) !== input.expectedSourceSha256) refuse("The history's source is not the current file.");
  if (evidence.resourceId !== input.resourceId || NOT_RESTORABLE.has(evidence.originSchemaVersion) ||
    (input.direction !== "undo" && input.direction !== "redo")) refuse("This edit cannot be undone by writing its bytes back.");
  const replaced = restoreBytes(evidence.replaced), written = restoreBytes(evidence.written);
  const current = await inspectMraiNativeContentStream(input.sourceBytes, {
    objectNumber: evidence.target.objectNumber, generation: evidence.target.generation,
  });
  const [spec, wanted, desired] = input.direction === "undo"
    ? [evidence.undo, written, replaced] as const
    : [evidence.redo, replaced, written] as const;
  const place = await portableWindow(spec, wanted, current.streamBytes);
  if (!place) refuse("The edit's bytes are no longer in the stream in one place; undo the later edits first.");
  const at = current.streamBytes.slice(place[0], place[1]);
  if (!equalBytes(at, wanted) || desired.length === 0 || place[1] <= place[0])
    refuse("The edit's bytes are no longer where it wrote them.");
  const structActualText = await planMraiNativeStructActualTextHistory(
    input.sourceBytes, (evidence as { structActualText?: unknown }).structActualText, input.direction);
  return {
    request: {
      identity: current.identity,
      expectedObjectSha256: current.objectSha256,
      expectedStreamSha256: current.streamSha256,
      operatorSpan: { start: place[0], end: place[1], expectedSha256: await sha256PdfBytes(at) },
      replacement: desired,
      untouchedObjects: [],
      ...(structActualText ? { structActualText } : {}),
    },
    history: {
      schemaVersion: HISTORY_SCHEMA_VERSION,
      originTransactionId: evidence.id,
      originSourceSha256: evidence.originSourceSha256,
      originCandidateSha256: evidence.originCandidateSha256,
      direction: input.direction,
    },
  };
}

export async function planMraiNativeByteRestoreHistory(input: PdfNativeByteRestoreHistoryInput): Promise<PdfNativeByteRestoreHistoryPlan> {
  if (isPdfNativeRestoreEvidence(input.originRecord)) return planPortableByteRestore(input, input.originRecord);
  const origin = input.originRecord;
  const source = record(origin.source), derived = record(origin.derivedCandidate), target = record(origin.target);
  if (await sha256PdfBytes(input.sourceBytes) !== input.expectedSourceSha256) refuse("The history's source is not the current file.");
  if (!mraiNativeRecordIsByteRestorable(origin) || source?.resourceId !== input.resourceId || !target ||
    (input.direction !== "undo" && input.direction !== "redo")) refuse("This edit cannot be undone by writing its bytes back.");
  const sourceLength = int(derived!.sourceLength)!, byteLength = int(derived!.byteLength)!;
  if (!(0 < sourceLength && sourceLength < byteLength && byteLength <= input.sourceBytes.length)) refuse("The edit is not part of this file's history.");
  const before = input.sourceBytes.slice(0, sourceLength);
  const after = input.sourceBytes.slice(0, byteLength);
  if (await sha256PdfBytes(before) !== source!.sha256 || await sha256PdfBytes(after) !== derived!.sha256)
    refuse("The edit is not part of this file's history.");
  const sourceStream = record(origin.pageContentConsolidation)
    ? await joinedSourceStream(before, origin)
    : (await inspectMraiNativeContentStream(before, { objectNumber: Number(target.objectNumber), generation: Number(target.generation) })).streamBytes;
  const envelope = await forwardEditEnvelope(origin, sourceStream);
  // The stream the forward wrote: the one object of its revision that holds
  // exactly its splice (the page's own stream, a Form's clone, a joined page).
  const owners: { objectNumber: number; generation: number }[] = [];
  for (const ref of objectsWrittenAfter(after, sourceLength)) {
    try {
      const written = await inspectMraiNativeContentStream(after, ref);
      if (equalBytes(written.streamBytes, envelope.writtenStream)) owners.push(ref);
    } catch (error) {
      if (!(error instanceof PdfNativeTransactionError)) throw error;
    }
  }
  if (owners.length !== 1) refuse("The stream the edit wrote cannot be found in its revision.");
  const current = await inspectMraiNativeContentStream(input.sourceBytes, owners[0]);
  const [reference, window, desired] = input.direction === "undo"
    ? [envelope.writtenStream, envelope.writtenWindow, sourceStream.slice(...envelope.sourceWindow)] as const
    : [sourceStream, envelope.sourceWindow, envelope.writtenStream.slice(...envelope.writtenWindow)] as const;
  const place = currentWindow(reference, window, current.streamBytes);
  if (!place) refuse("The edit's bytes are no longer in the stream in one place; undo the later edits first.");
  const wanted = current.streamBytes.slice(place[0], place[1]);
  if (!equalBytes(wanted, reference.slice(window[0], window[1])) || desired.length === 0 || place[1] <= place[0])
    refuse("The edit's bytes are no longer where it wrote them.");
  // The structure elements it wrote beside its splice go back with it.
  const structActualText = await planMraiNativeStructActualTextHistory(input.sourceBytes, origin.structActualText, input.direction);
  return {
    request: {
      identity: current.identity,
      expectedObjectSha256: current.objectSha256,
      expectedStreamSha256: current.streamSha256,
      operatorSpan: { start: place[0], end: place[1], expectedSha256: await sha256PdfBytes(wanted) },
      replacement: desired,
      untouchedObjects: [],
      ...(structActualText ? { structActualText } : {}),
    },
    history: {
      schemaVersion: HISTORY_SCHEMA_VERSION,
      originTransactionId: String(origin.id),
      originSourceSha256: String(source!.sha256),
      originCandidateSha256: String(derived!.sha256),
      direction: input.direction,
    },
  };
}
