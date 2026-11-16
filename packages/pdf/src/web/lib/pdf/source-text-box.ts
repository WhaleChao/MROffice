/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfNativeListTextSelectionProof } from "./engine/native/list-text-fidelity";
import type { PdfFallbackTextPiece, PdfFallbackTextProofData } from "./source-text-fallback";
import { scanPdfSourceTextBoxAppearanceSeals, type PdfSourceTextBoxSealIndex } from "./source-text-box-seal-scan";
import { boxRectContains, expandRectForGlyphCover, type BoxRect } from "@mrnote/core/web/lib/pdf/text-box-geometry.ts";
import { startResilientWorker } from "@mrnote/core/web/lib/worker-start.ts";

function readFallbackPieces(value: unknown): readonly PdfFallbackTextPiece[] | null {
  if (!Array.isArray(value) || value.length < 1 || value.length > 1024) return null;
  const pieces: PdfFallbackTextPiece[] = [];
  for (const entry of value) {
    if (!entry || typeof entry !== "object") return null;
    const piece = entry as Record<string, unknown>;
    const ints = [piece.object, piece.generation, piece.start, piece.end];
    if (!ints.every((number) => Number.isSafeInteger(number)) || (piece.object as number) < 1 ||
      (piece.generation as number) < 0 || (piece.generation as number) > 65535 || (piece.start as number) < 0 ||
      (piece.end as number) <= (piece.start as number) || typeof piece.sha256 !== "string" ||
      !/^[a-f0-9]{64}$/.test(piece.sha256)) return null;
    const ranges = piece.codeRanges;
    if (ranges !== undefined) {
      if (!Array.isArray(ranges) || !ranges.length || ![1, 2, 3, 4].includes(piece.codeWidth as number) ||
        typeof piece.operator !== "string" || !/^[A-Za-z0-9+/]+={0,2}$/.test(piece.operator)) return null;
      let previous = 0;
      for (const pair of ranges) {
        if (!Array.isArray(pair) || pair.length !== 2 || !pair.every(Number.isSafeInteger) ||
          pair[0] < previous || pair[0] >= pair[1]) return null;
        previous = pair[1];
      }
    } else if (piece.codeWidth !== undefined || piece.operator !== undefined) return null;
    const next = piece.continuation as Record<string, unknown> | undefined;
    if (next !== undefined && (ranges !== undefined || !next || typeof next !== "object" ||
      Object.keys(next).length !== 4 || ![next.object, next.generation, next.end].every(Number.isSafeInteger) ||
      (next.object as number) < 1 || (next.end as number) < 1 || typeof next.sha256 !== "string" ||
      !/^[a-f0-9]{64}$/.test(next.sha256))) return null;
    pieces.push(Object.freeze({
      ...(next ? { continuation: Object.freeze({ object: next.object as number, generation: next.generation as number,
        end: next.end as number, sha256: next.sha256 as string }) } : {}),
      object: piece.object as number, generation: piece.generation as number,
      start: piece.start as number, end: piece.end as number, sha256: piece.sha256,
      ...(ranges ? { codeRanges: (ranges as [number, number][]).map(([a, b]) => [a, b] as const),
        codeWidth: piece.codeWidth as number, operator: piece.operator as string } : {}),
    }));
  }
  return Object.freeze(pieces);
}

/** A span binding's durable proof: what re-opens and identifies the box,
 * never a native write grant (and small enough for the directive token). */
export function spanBindingProof(proof: PdfNativeListTextSelectionProof, data: PdfFallbackTextProofData): PdfNativeListTextSelectionProof {
  const evidence = proof.evidence;
  const first = data.pieces[0];
  const renderingMode = sourceRenderingModeOf(data, evidence);
  return Object.freeze({
    target: Object.freeze({ objectNumber: first.object, generation: first.generation }),
    evidence: Object.freeze({
      schemaVersion: "mrai.pdf-source-text-fallback.v1",
      layoutKind: evidence.layoutKind === "list-body" ? "list-body" : "paragraph-line",
      originalText: evidence.originalDisplayText,
      originalDisplayText: evidence.originalDisplayText,
      // The box's identity (page, first item, text); its reader items are
      // read again from the page and would only fill the directive token.
      selection: evidence.selection ? Object.freeze({ ...evidence.selection, items: undefined }) : null,
      fontSize: evidence.fontSize,
      lineSpacing: evidence.lineSpacing,
      textOperatorSpan: Object.freeze({ start: first.start, end: first.end, sha256: first.sha256 }),
      // The run's advance, with its appearance below: what the rules drawn
      // beside the hidden run are found by when the box is deleted
      // (PdfPageWindow commitNativeTextObjectDelete, hidden source).
      ...(typeof evidence.originalAdvance === "number" ? { originalAdvance: evidence.originalAdvance } : {}),
      // How the source drew the run (`Tr`): 3 or 7 is a run that paints
      // nothing, the ink the reader sees being a scan under it; the box then
      // covers that ink (pdfSourceTextBoxCoversInk). Read where the source
      
      ...(renderingMode !== undefined ? { renderingMode } : {}),
    }) as unknown as PdfNativeListTextSelectionProof["evidence"],
    properties: null,
    ...(proof.sourceAppearance ? { sourceAppearance: proof.sourceAppearance } : {}),
  });
}

function sourceRenderingModeOf(data: PdfFallbackTextProofData, evidence: object): number | undefined {
  const mode = data.renderingMode ?? (evidence as { renderingMode?: unknown }).renderingMode;
  return typeof mode === "number" && Number.isInteger(mode) && mode >= 0 && mode <= 7 ? mode : undefined;
}

/** Whether a box bound to source text covers what is under it: its source
 * run paints nothing (`3 Tr`/`7 Tr`, an OCR layer over a scanned page), so
 * hiding the run hides no ink, and the scanned glyphs would show through a
 * clear box (user audit PDF 0107). Such a box is drawn over an opaque fill
 * in the colour of the page around the line, as Acrobat's edit of OCR text
 * does; a box over a run that paints its own glyphs stays clear (the run is
 * hidden in the bytes). Only a span binding (version 3) carries it; the
 * server checks it against the mode the page draws the pieces in. */
export function pdfSourceTextBoxCoversInk(
  binding: Readonly<{ version?: number; originalProof?: { evidence?: unknown } }> | null | undefined,
): boolean {
  if (binding?.version !== 3) return false;
  const mode = (binding.originalProof?.evidence as { renderingMode?: unknown } | undefined)?.renderingMode;
  return mode === 3 || mode === 7;
}

/** How far a frame may stand inside the line it replaced and still hold it
 * (page fractions, under half a point): a frame fitted to its rows again. */
const INK_COVER_HOLD_EPSILON = 0.000_5;

/** The stationary cover of a box over a run that paints nothing (PDF 0107,
 * pdfSourceTextBoxCoversInk) once its frame no longer holds the line it
 * replaced -- dragged away, or set narrower than the scanned words: the
 * scanned letters stay where the scan has them, so the page's colour stays
 * there too (the line's rect, `origin*`, with the glyph bleed of every other
 * cover: expandRectForGlyphCover), and the box itself is clear wherever it
 * stands, so the scan under its new place shows between its glyphs -- as
 * WPS's edit of a converted scan does. Null while the frame holds the whole
 * line: then the box is drawn on its own fill (one cover, never two).
 * The writer paints this rect inside the same annotation (its rect is the
 * union of the two); the live overlay, the editing frame, the flattened
 * export and the raster appearance draw the same. Rounded as the record
 * writes it, so every reader of a box draws one rect. */
export function pdfSourceTextBoxInkCoverRect(edit: Readonly<{
  sourceText?: Readonly<{ version?: number; originalProof?: { evidence?: unknown } }> | null;
  backgroundColor?: unknown; fontSize?: unknown; angle?: unknown;
  x?: unknown; y?: unknown; width?: unknown; height?: unknown;
  originX?: unknown; originY?: unknown; originWidth?: unknown; originHeight?: unknown;
}> | null | undefined): BoxRect | null {
  if (!edit || !pdfSourceTextBoxCoversInk(edit.sourceText) ||
    typeof edit.backgroundColor !== "string" || !/^#[0-9a-f]{6}$/i.test(edit.backgroundColor.trim())) return null;
  const rect = (x: unknown, y: unknown, width: unknown, height: unknown): BoxRect | null => {
    const values = [x, y, width, height].map(Number);
    return values.every(Number.isFinite) && values[2] > 0 && values[3] > 0
      ? { x: values[0], y: values[1], width: values[2], height: values[3] } : null;
  };
  const origin = rect(edit.originX, edit.originY, edit.originWidth, edit.originHeight);
  const box = rect(edit.x, edit.y, edit.width, edit.height);
  return origin && box ? pdfInkCoverRectOutsideFrame(box, origin, edit.fontSize, pdfTextBoxQuarterTurned(edit.angle)) : null;
}

/** Whether a box's angle is a quarter turn (a line on a page shown turned). */
export function pdfTextBoxQuarterTurned(angle: unknown): boolean {
  const degrees = Number(angle);
  if (!Number.isFinite(degrees)) return false;
  const turn = ((degrees % 360) + 360) % 360;
  return Math.abs(turn - 90) < 1e-6 || Math.abs(turn - 180) < 1e-6 || Math.abs(turn - 270) < 1e-6;
}

/** The cover pdfSourceTextBoxInkCoverRect names, from the frame where it
 * stands and the line it replaced (`origin`): null while the frame holds
 * the line. The editing frame asks it before the box has a binding. */
export function pdfInkCoverRectOutsideFrame(
  frame: BoxRect,
  origin: BoxRect,
  fontSizeFraction?: unknown,
  






  turned = false,
): BoxRect | null {
  if (boxRectContains(frame, origin, INK_COVER_HOLD_EPSILON)) return null;
  const fontSize = Number(fontSizeFraction);
  const size = Number.isFinite(fontSize) && fontSize > 0 ? fontSize : undefined;
  // The frame's row is top-aligned with the glyph band where the line was
  // (the frame is recorded at the band's own top), wherever it stands now.
  const band = turned && frame.height > origin.height ? { ...origin, height: frame.height } : origin;
  const expanded = expandRectForGlyphCover(band, size);
  const cover = turned
    ? { ...expanded, x: band.x + band.width / 2 - expanded.width / 2 }
    : expanded;
  const round = (value: number) => Number(value.toFixed(6));
  const rounded = { x: round(cover.x), y: round(cover.y), width: round(cover.width), height: round(cover.height) };
  return rounded.width > 0.000_001 && rounded.height > 0.000_001 &&
    rounded.x + rounded.width <= 1.000_001 && rounded.y + rounded.height <= 1.000_001 ? rounded : null;
}

export const normalizedText = (text: string) => text.replace(/\s+/gu, " ").trim();

// Keep the transport within the existing server's decoded source-proof budget.
export const MAX_PDF_SOURCE_TEXT_BOX_TOKEN_CHARS = 50_000;

/** Glyph quads are a disposable rendering cache, not part of the source-show
 * witness. Keep native operators/resources/property proofs intact in durable
 * recipes; current display geometry is acquired from the loaded PDF. */
export function durableSourceTextProof(proof: PdfNativeListTextSelectionProof): PdfNativeListTextSelectionProof {
  return {
    target: proof.target, evidence: proof.evidence, properties: proof.properties,
    ...(proof.propertyRuns ? {propertyRuns: proof.propertyRuns} : {}),
    ...(proof.sourceAppearance ? {sourceAppearance: proof.sourceAppearance} : {}),
  };
}

export interface PdfSourceTextBoxBinding {
  

  readonly version: 1 | 2 | 3;
  /** Version 3: every span the box's text is drawn by, hidden by span. */
  readonly pieces?: readonly PdfFallbackTextPiece[];
  /** Existing PDF owner identity, retained only after a portable import:
   * always on version 2, and on a version 3 binding read back from the file
   * it was written in (the owner is then named as the file states it, not
   * by this host's resource and page). */
  readonly sourceIdentity?: string;
  readonly resourceId: string;
  readonly pageId: string;
  readonly boxId: string;
  readonly expectedAppearance: string | null;
  readonly page: number;
  readonly object: number;
  readonly generation: number;
  readonly operator: string;
  readonly sha256: string;
  /** Original read proof only. Current edits are managed-box writes, never glyph writes. */
  readonly originalProof: PdfNativeListTextSelectionProof;
}

/** Whether a box's binding names this reader item as the source line it
 * stands over: the same text, at the item index it was bound at or, when
 * that index moved, where the line stood. An edit above the line that
 * changes how many items the reader counts (a line set on two rows) moves
 * every index after it; the hidden line itself does not move, and a box
 * bound by its index alone was then opened as the hidden line under it and
 * saved over it again (refused: "source text box identity not proven"). */
export function pdfSourceTextBindingNamesItem(
  binding: Readonly<{ originalProof?: { evidence?: { selection?: unknown } } }> | null | undefined,
  pageNumber: number,
  item: Readonly<{ itemIndex?: number; str: string; transform?: readonly number[] }>,
): boolean {
  const origin = binding?.originalProof?.evidence?.selection as
    | Readonly<{ pageNumber?: number; item?: Readonly<{ itemIndex?: number; str?: string; transform?: readonly number[] }> }>
    | null
    | undefined;
  // As the API stores the binding: a glyph whose ToUnicode is U+0000 (a
  // bullet's) is kept there as U+FFFD, and read from the page as U+0000. A
  // box saved over such a line refused to open again after a reload, and
  
  const flat = (text: string) => text.replace(/\s+/gu, "").replaceAll("\u0000", "\ufffd");
  if (!origin?.item || origin.pageNumber !== pageNumber || typeof origin.item.str !== "string") return false;
  const bound = flat(origin.item.str);
  const read = flat(item.str);
  // The hidden run is wrapped in a marked-content span once saved (its
  // ActualText carries the box's words for other readers), and the reader
  // ends an item at the span's end: a line it had read as one item with
  // what followed it on the same baseline is now read up to the run alone.
  // Such a cut item is the line when it stands where the line stood.
  const cut = bound !== read && read.length > 0 && bound.startsWith(read);
  if (bound !== read && !cut) return false;
  if (!cut && origin.item.itemIndex === item.itemIndex) return true;
  const from = origin.item.transform;
  const at = item.transform;
  return Boolean(from && at && from.length === 6 && at.length === 6) &&
    Math.abs(Number(from![4]) - Number(at![4])) < 0.01 &&
    Math.abs(Number(from![5]) - Number(at![5])) < 0.01;
}

export function readPdfSourceTextBoxBinding(
  value: unknown,
): PdfSourceTextBoxBinding | null {
  if (!value || typeof value !== "object") return null;
  const b = value as PdfSourceTextBoxBinding;
  const pieces = b.version === 3 ? readFallbackPieces(b.pieces) : null;
  if (
    (b.version !== 1 && b.version !== 2 && b.version !== 3) ||
    ((b.version === 2 || (b.version === 3 && b.sourceIdentity !== undefined)) &&
      (typeof b.sourceIdentity !== "string" ||
        !/^[a-f0-9]{64}$/.test(b.sourceIdentity) || b.expectedAppearance === null)) ||
    (b.version === 1 && b.sourceIdentity !== undefined) ||
    (b.version === 3 && (!pieces || pieces[0].object !== b.object ||
      pieces[0].generation !== b.generation || pieces[0].sha256 !== b.sha256)) ||
    (b.version !== 3 && b.pieces !== undefined) ||
    typeof b.resourceId !== "string" ||
    !b.resourceId ||
    typeof b.pageId !== "string" ||
    !b.pageId ||
    typeof b.boxId !== "string" ||
    !b.boxId ||
    b.boxId.length > 120 ||
    (b.expectedAppearance !== null &&
      (typeof b.expectedAppearance !== "string" ||
        !/^[a-f0-9]{64}$/.test(b.expectedAppearance))) ||
    !Number.isSafeInteger(b.page) ||
    b.page < 1 ||
    !Number.isSafeInteger(b.object) ||
    b.object < 1 ||
    !Number.isSafeInteger(b.generation) ||
    b.generation < 0 ||
    b.generation > 65535 ||
    typeof b.operator !== "string" ||
    b.operator.length < 4 ||
    b.operator.length > 32768 ||
    !/^[A-Za-z0-9+/]+={0,2}$/.test(b.operator) ||
    !/^[a-f0-9]{64}$/.test(b.sha256) ||
    b.originalProof?.target?.objectNumber !== b.object ||
    b.originalProof.target.generation !== b.generation ||
    typeof b.originalProof.evidence?.originalDisplayText !== "string"
  )
    return null;
  return {
    version: b.version,
    ...(b.sourceIdentity !== undefined ? { sourceIdentity: b.sourceIdentity } : {}),
    ...(pieces ? { pieces } : {}),
    resourceId: b.resourceId,
    pageId: b.pageId,
    boxId: b.boxId,
    expectedAppearance: b.expectedAppearance,
    page: b.page,
    object: b.object,
    generation: b.generation,
    operator: b.operator,
    sha256: b.sha256,
    originalProof: durableSourceTextProof(b.originalProof),
  };
}

/** Source-file metadata carries no host scope. Bind it to the current import;
 * this is read evidence only, never permission to write a different resource. */
export function readPortablePdfSourceTextBoxBinding(
  value: unknown,
  scope: { resourceId: string; pageId: string },
  expectedAppearance: string,
): PdfSourceTextBoxBinding | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const portable = value as Record<string, unknown>;
  
  // version 3) is recorded with the spans it hides. Read without them it
  // was no record at all: every box of its page was dropped on a reload,
  // the hidden line was offered as a line of the page, and set again it was
  // written a second time under the box still drawn over it (Attention
  
  const spans = "pieces" in portable;
  const keys = ["schemaVersion", "sourceIdentity", "boxId", "page", "object",
    "generation", "operator", "sha256", "originalProof", ...(spans ? ["pieces"] : [])];
  if (portable.schemaVersion !== "mrai.pdf-source-text-box.portable.v1" ||
    Object.keys(portable).length !== keys.length || keys.some(key => !(key in portable))) return null;
  return readPdfSourceTextBoxBinding({
    version: spans ? 3 : 2,
    ...(spans ? { pieces: portable.pieces } : {}),
    sourceIdentity: portable.sourceIdentity,
    resourceId: scope.resourceId,
    pageId: scope.pageId,
    boxId: portable.boxId,
    page: portable.page,
    object: portable.object,
    generation: portable.generation,
    operator: portable.operator,
    sha256: portable.sha256,
    originalProof: portable.originalProof,
    expectedAppearance,
  });
}

/** Whether a binding was read back from its file (a portable import): the
 * page's own drawing already paints its box, and the server names its owner
 * by the identity the file states. */
export function pdfSourceTextBindingIsPortable(
  binding: Readonly<{ version?: number; sourceIdentity?: string }> | null | undefined,
): boolean {
  return binding?.version === 2 || (binding?.version === 3 && typeof binding.sourceIdentity === "string");
}














export function pdfSourceTextBindingSealedForSave<Binding extends PdfSourceTextBoxBinding>(
  binding: Binding,
  newestSeal: string | null,
): Binding {
  if (pdfSourceTextBindingIsPortable(binding) && binding.expectedAppearance) return binding;
  return { ...binding, expectedAppearance: newestSeal };
}

export function pdfSourceTextBoxOperator(
  b: PdfSourceTextBoxBinding,
): Uint8Array {
  return Uint8Array.from(atob(b.operator), (character) =>
    character.charCodeAt(0),
  );
}

export function pdfSourceTextBoxHiddenOperator(
  b: PdfSourceTextBoxBinding,
): Uint8Array {
  const raw = atob(b.operator);
  return Uint8Array.from(`\n3 Tr\n${raw}\n0 Tr\n`, (character) =>
    character.charCodeAt(0),
  );
}

function sourceTextBoxAnnotationName(boxId: string): string {
  return `MRNote:textbox.${encodeURIComponent(boxId).replace(
    /[!'()*]/g,
    (character) => `%${character.charCodeAt(0).toString(16).toUpperCase()}`,
  )}`;
}

/** The seals `bindings` ask for, read from a whole-file index. Throws as the
 * direct read does: a missing page, an ambiguous box, a malformed seal. */
export function pdfSourceTextBoxSealsFromIndex(
  index: PdfSourceTextBoxSealIndex,
  bindings: readonly PdfSourceTextBoxBinding[],
): Map<string, string | null> {
  const seals = new Map<string, string | null>();
  for (const binding of bindings) {
    if (!(binding.page >= 1 && binding.page <= index.pageCount)) throw new Error("The source text-box page changed.");
    if (index.unreadablePages.includes(binding.page)) throw new Error("The source text-box page could not be read.");
    const entry = index.boxes[`${binding.page}\u0000${sourceTextBoxAnnotationName(binding.boxId)}`];
    if (entry?.invalid) throw new Error("The source text-box appearance is invalid.");
    if (entry && entry.matches > 1) throw new Error("The source text-box identity is ambiguous.");
    seals.set(binding.boxId, entry?.seal ?? null);
  }
  return seals;
}

export async function readPdfSourceTextBoxAppearanceSeals(
  sourceBytes: Uint8Array,
  bindings: readonly PdfSourceTextBoxBinding[],
): Promise<Map<string, string | null>> {
  return pdfSourceTextBoxSealsFromIndex(await scanPdfSourceTextBoxAppearanceSeals(sourceBytes), bindings);
}

// Seal indexes by source SHA-256, read off the main thread (the worker
// checks the bytes against the digest first). A text-box save reads its
// seals from here instead of loading the whole PDF on the reader's thread
// while the editor waits to let go (~0.3-0.5 s for a 4 MB file).
const sealIndexes = new Map<string, Promise<PdfSourceTextBoxSealIndex | null>>();
let sealWorker: Worker | null = null;
let sealRequestId = 0;
const sealRequests = new Map<number, (index: PdfSourceTextBoxSealIndex | null) => void>();
const sealWorkerStart = startResilientWorker("pdf-text-box-seals");

function sealIndexWorker(): Worker | null {
  if (sealWorker) return sealWorker;
  if (typeof Worker === "undefined") return null;
  // One more start after a pause when one fails (lib/worker-start.ts); none:
  // no seal index, as with no Worker.
  const started = sealWorkerStart.create(() => new Worker(new URL("./source-text-box-seals.worker.ts", import.meta.url),
    { type: "module", name: "mrnote-pdf-text-box-seals" }));
  if (!started) return null;
  sealWorker = started;
  sealWorker.onmessage = (event: MessageEvent<{ id: number; index?: PdfSourceTextBoxSealIndex }>) => {
    sealWorkerStart.started(started);
    const resolve = sealRequests.get(event.data.id);
    sealRequests.delete(event.data.id);
    resolve?.(event.data.index ?? null);
  };
  sealWorker.onerror = () => {
    sealWorkerStart.failed(started, "error");
    for (const resolve of sealRequests.values()) resolve(null);
    sealRequests.clear();
    sealWorker?.terminate();
    sealWorker = null;
  };
  return sealWorker;
}

/** Reads the seal index of the source `sha256` names, off the main thread,
 * once per source. Null where no worker runs, or the bytes are not that
 * source. `bytes` is copied, never kept. */
export function warmPdfSourceTextBoxAppearanceSeals(
  sha256: string,
  bytes: Uint8Array,
): Promise<PdfSourceTextBoxSealIndex | null> {
  const known = sealIndexes.get(sha256);
  if (known) return known;
  const worker = sealIndexWorker();
  if (!worker || !/^[a-f0-9]{64}$/.test(sha256) || !bytes.length) return Promise.resolve(null);
  const id = (sealRequestId += 1);
  const copy = Uint8Array.from(bytes);
  const pending = new Promise<PdfSourceTextBoxSealIndex | null>((resolve) => {
    sealRequests.set(id, resolve);
    // A worker that never answers (one that failed to load without an
    // error event) is given up on: the save reads the bytes itself.
    setTimeout(() => {
      if (sealRequests.get(id) !== resolve) return;
      sealRequests.delete(id);
      resolve(null);
    }, 10_000);
    try {
      worker.postMessage({ id, sha256, bytes: copy.buffer }, [copy.buffer]);
    } catch {
      sealRequests.delete(id);
      resolve(null);
    }
  }).then((index) => {
    if (!index) sealIndexes.delete(sha256);
    return index;
  });
  sealIndexes.set(sha256, pending);
  // The latest few sources only (each save makes a new one).
  while (sealIndexes.size > 3) sealIndexes.delete(sealIndexes.keys().next().value!);
  return pending;
}

/** The seals for `bindings` from the source's index when it was read (or is
 * being read), else null: the caller reads them from the bytes itself. */
export async function cachedPdfSourceTextBoxAppearanceSeals(
  sha256: string,
  bindings: readonly PdfSourceTextBoxBinding[],
): Promise<Map<string, string | null> | null> {
  const pending = sealIndexes.get(sha256);
  if (!pending) return null;
  // Still being read: waited for briefly, never longer than reading the
  // bytes here would take.
  const index = await Promise.race([pending,
    new Promise<null>((resolve) => setTimeout(() => resolve(null), 1_500))]);
  return index ? pdfSourceTextBoxSealsFromIndex(index, bindings) : null;
}

export function pdfSourceTextBoxLineHeightRatio(binding: PdfSourceTextBoxBinding | undefined): number | undefined {
  const evidence = binding?.originalProof.evidence;
  if (!evidence || evidence.layoutKind !== "paragraph-line") return undefined;
  const leading = Math.abs(evidence.lineSpacing);
  const size = evidence.fontSize;
  if (!Number.isFinite(leading) || !Number.isFinite(size) || leading <= 0 || size <= 0) return undefined;
  const ratio = leading / size;
  return ratio >= 0.5 && ratio <= 10 ? ratio : undefined;
}
