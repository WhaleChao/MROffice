/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure, dependency-free helpers for two PowerPoint / WPS "Home"-tab gestures on
 * presentation text that are worth unit-testing in isolation:
 *
 *  1. CHANGE CASE (the "Aa" button + Shift+F3): UPPERCASE / lowercase /
 *     Capitalize Each Word / Sentence case / tOGGLE over the selected run text.
 *     The pure string logic is PORTED from the native-document editor's
 *     `documentTextTransforms.ts` (`ChangeCaseMode` + `transformCase` +
 *     `nextCaseModeForShiftF3`) so the two surfaces apply IDENTICAL casing — the
 *     same Unicode-aware Title Case, the same sentence detector (incl. CJK
 *     terminators 。！？), the same Shift+F3 three-state rotation.
 *     {@link applyCaseToRuns} then re-applies that pure transform across the
 *     presentation `TextRun[]` model (selection-wide, mark-preserving),
 *     mirroring how `presentation-runs.applyRunFormat` walks runs without
 *     treating a style boundary as a word or sentence boundary.
 *
 *  2. LOCK OBJECT (the per-block `locked` flag, PowerPoint's Lock): a locked
 *     block can't be moved / resized / rotated (nor grouped, reshaped or
 *     cropped); its text, format and Delete stay open. {@link
 *     canManipulateBlock} is what the frame gestures read, {@link
 *     canEditBlockContent} what content, format and Delete read.
 *
 * SERIALIZATION: the case transform only rewrites the CHARACTERS of existing
 * runs (never adds a node / style / attribute), and `locked` is an additive,
 * absent-by-default boolean — a deck that never invokes either path serializes
 * byte-for-byte unchanged.
 *
 * No React / no DOM here (the component owns the selection + commit), so every
 * function below is pure and inputs are never mutated.
 */

import {
  mergeRuns,
  textFromRuns,
  type RunStyle,
  type TextRun,
} from "./presentation-point-runs";
import { officeSentenceCase } from "./office-text-case";

// ===========================================================================
// CHANGE CASE — pure string transforms (ported from documentTextTransforms.ts)
// ===========================================================================

export type ChangeCaseMode =
  "upper" | "lower" | "title" | "sentence" | "toggle";

/** Ordered list of the five modes (stable for menus and tests). */
export const CHANGE_CASE_MODES: readonly ChangeCaseMode[] = [
  "sentence",
  "upper",
  "lower",
  "title",
  "toggle",
];

/**
 * The Shift+F3 cycle, matching Word's three-state rotation: Sentence case →
 * UPPERCASE → lowercase → (back to Sentence case). Title Case and tOGGLE are
 * reachable from the menu but are NOT part of the keyboard cycle (Word doesn't
 * cycle through them either). PURE.
 */
const SHIFT_F3_CYCLE: readonly ChangeCaseMode[] = [
  "sentence",
  "upper",
  "lower",
];

/**
 * Given the text currently in the selection, decide which case mode the next
 * Shift+F3 press should apply. We look at what the text already looks like and
 * advance to the next state in the cycle, so repeated presses rotate
 * Sentence → UPPER → lower → Sentence the way Word does. When the text matches
 * none of the cycle's outputs (e.g. mixed/Title case), we start the cycle at
 * UPPERCASE — the same first transition Word makes from arbitrary text. PURE.
 */
export function nextCaseModeForShiftF3(text: string): ChangeCaseMode {
  for (let i = 0; i < SHIFT_F3_CYCLE.length; i += 1) {
    const mode = SHIFT_F3_CYCLE[i];
    if (transformCase(text, mode) === text) {
      return SHIFT_F3_CYCLE[(i + 1) % SHIFT_F3_CYCLE.length];
    }
  }
  return "upper";
}

/**
 * Title Case: upper-case the first letter of every "word" (a maximal run of
 * letters/digits/apostrophes) and lower-case the rest of the word. Uses Unicode
 * letter classes so non-ASCII letters (accented Latin, Greek, Cyrillic, …)
 * capitalize correctly; CJK has no case and is left unchanged by the underlying
 * `toUpperCase`/`toLowerCase`. PURE.
 */
function toTitleCase(text: string): string {
  return text.replace(/[\p{L}\p{M}\p{N}']+/gu, (word) => {
    // The first LETTER, not the first UTF-16 unit: a supplementary-plane
    // letter (Deseret, Gothic) is a surrogate pair, and half a pair has no
    // case (PPT-B251).
    const head = String.fromCodePoint(word.codePointAt(0) ?? 0);
    const first = head.toUpperCase();
    const rest = word.slice(head.length).toLowerCase();
    return first + rest;
  });
}

/**
 * Sentence case: lower-case everything, then upper-case the first letter of
 * each sentence. A "sentence start" is the first letter of the string or the
 * first letter following sentence-ending punctuation (`.`, `!`, `?`, and their
 * full-width CJK forms 。！？) plus any whitespace/closing quotes. PURE.
 */
function toSentenceCase(text: string): string {
  return officeSentenceCase(text);
}

/** tOGGLE case: invert the case of every cased character. PURE. */
function toToggleCase(text: string): string {
  let out = "";
  for (const ch of text) {
    out += toggleCaseCharacter(ch);
  }
  return out;
}

function toggleCaseCharacter(character: string): string {
  const upper = character.toUpperCase();
  const lower = character.toLowerCase();
  if (character === upper && character !== lower) return lower;
  if (character === lower && character !== upper) return upper;
  return character;
}

/**
 * Apply a {@link ChangeCaseMode} to a plain string. PURE and the single source
 * of truth for the case logic — both the unit tests and the run-applying
 * function below route through it.
 */
export function transformCase(text: string, mode: ChangeCaseMode): string {
  switch (mode) {
    case "upper":
      return text.toUpperCase();
    case "lower":
      return text.toLowerCase();
    case "title":
      return toTitleCase(text);
    case "sentence":
      return toSentenceCase(text);
    case "toggle":
      return toToggleCase(text);
    default:
      return text;
  }
}

// ===========================================================================
// CHANGE CASE — apply across the presentation TextRun[] model
// ===========================================================================

/** Carry a run's style forward as a fresh copy (never share the object). */
function cloneRunStyle(style: RunStyle | undefined): RunStyle | undefined {
  return style ? { ...style } : undefined;
}

const CASE_WORD_CHARACTER_RE = /^[\p{L}\p{M}\p{N}']$/u;
const CASE_LETTER_RE = /^\p{L}$/u;
const SENTENCE_TERMINATOR_RE = /^[.!?。！？]$/u;
const SENTENCE_LEAD_CHARACTER_RE = /^[\s'")\]]$/u;

function sentenceCharacterOutputLength(
  character: string,
  uppercaseFirstLetter: boolean,
): number {
  const lowered = character.toLowerCase();
  if (!uppercaseFirstLetter) return lowered.length;
  const first = Array.from(lowered)[0] ?? "";
  return first.toUpperCase().length + lowered.length - first.length;
}

function caseCharacterOutputLength(
  character: string,
  mode: ChangeCaseMode,
  options: {
    titleWordStart: boolean;
    sentenceStart: boolean;
  },
): number {
  switch (mode) {
    case "upper":
      return character.toUpperCase().length;
    case "lower":
      return character.toLowerCase().length;
    case "title":
      return (
        options.titleWordStart
          ? character.toUpperCase()
          : character.toLowerCase()
      ).length;
    case "sentence":
      return sentenceCharacterOutputLength(character, options.sentenceStart);
    case "toggle":
      return toggleCaseCharacter(character).length;
    default:
      return character.length;
  }
}

/**
 * Transform selected run slices as one logical string while retaining a slice
 * for each original style run. The logical source is transformed exactly once;
 * a separate, out-of-band ownership scan measures how many output code units
 * belong to every source slice. This keeps Unicode context intact (notably the
 * Greek final sigma) while assigning expanding mappings such as `ß -> SS` and
 * `ﬁ -> FI` to the style that owned the source character.
 *
 * The ownership scan advances through source characters and run boundaries
 * monotonically. Its storage is bounded by the run count and its work is
 * O(source length + output length + run count); it never constructs a
 * source-length intermediate for every boundary.
 */
function transformCaseAcrossSlices(
  slices: readonly string[],
  mode: ChangeCaseMode,
): string[] {
  const source = slices.join("");
  if (!source || slices.length === 0) return [];
  const transformed = transformCase(source, mode);
  if (slices.length === 1) return [transformed];

  const outputLengths = slices.map(() => 0);
  let sliceIndex = 0;
  let sliceEnd = slices[0].length;
  let sourceOffset = 0;
  let insideTitleWord = false;
  let afterSentenceTerminator = false;

  for (const character of source) {
    while (sliceIndex < slices.length - 1 && sourceOffset >= sliceEnd) {
      sliceIndex += 1;
      sliceEnd += slices[sliceIndex].length;
    }

    const isWordCharacter = CASE_WORD_CHARACTER_RE.test(character);
    const isLetter = CASE_LETTER_RE.test(character);
    const titleWordStart = isWordCharacter && !insideTitleWord;
    const sentenceStart =
      isLetter && (sourceOffset === 0 || afterSentenceTerminator);
    outputLengths[sliceIndex] += caseCharacterOutputLength(character, mode, {
      titleWordStart,
      sentenceStart,
    });

    insideTitleWord = isWordCharacter;
    if (SENTENCE_TERMINATOR_RE.test(character)) {
      afterSentenceTerminator = true;
    } else if (afterSentenceTerminator) {
      if (isLetter || !SENTENCE_LEAD_CHARACTER_RE.test(character)) {
        afterSentenceTerminator = false;
      }
    }
    sourceOffset += character.length;
  }

  const measuredOutputLength = outputLengths.reduce(
    (sum, length) => sum + length,
    0,
  );
  if (measuredOutputLength !== transformed.length) {
    throw new Error("Change-case output ownership could not be authenticated.");
  }

  const result: string[] = [];
  let outputOffset = 0;
  for (const length of outputLengths) {
    result.push(transformed.slice(outputOffset, outputOffset + length));
    outputOffset += length;
  }
  return result;
}

/**
 * Rewrite the CHARACTERS of the character range `[start, end)` of a runs array
 * by applying {@link transformCase}, splitting runs at the selection boundaries
 * and transforming only the covered slice of each run, then merging
 * adjacent-equal runs. Mirrors `presentation-runs.applyRunFormat` (same
 * clamp + per-run walk) but it changes TEXT, not style — every run keeps its own
 * style.
 *
 * Covered slices are transformed as ONE logical selection. Run boundaries are
 * retained only to restore styles; they never invent sentence/word starts.
 * A bounded out-of-band ownership map keeps the mapping correct when Unicode
 * case maps change length (`ß → SS`, `ﬁ → FI`) without inserting characters
 * that could alter contextual casing.
 *
 * Returns a fresh, normalized array. An out-of-range / reversed selection is a
 * no-op (returns the merged input). PURE — the input is not mutated.
 *
 * NOTE: the flat block `text` (the runs⟷text invariant) stays in sync because
 * the caller recomputes it via `textFromRuns(result)`.
 */
export function applyCaseToRuns(
  runs: readonly TextRun[],
  start: number,
  end: number,
  mode: ChangeCaseMode,
): TextRun[] {
  const base = mergeRuns(runs);
  const total = textFromRuns(base).length;
  const from = Math.min(Math.max(Math.round(start), 0), total);
  const to = Math.min(Math.max(Math.round(end), 0), total);
  if (from >= to) return base.length ? base : [];

  const coveredSlices: string[] = [];
  let scanCursor = 0;
  for (const run of base) {
    const runStart = scanCursor;
    const runEnd = scanCursor + run.text.length;
    scanCursor = runEnd;
    if (runEnd <= from || runStart >= to) continue;
    const localFrom = Math.max(from, runStart) - runStart;
    const localTo = Math.min(to, runEnd) - runStart;
    coveredSlices.push(run.text.slice(localFrom, localTo));
  }
  const transformedSlices = transformCaseAcrossSlices(coveredSlices, mode);

  const out: TextRun[] = [];
  let cursor = 0;
  let transformedIndex = 0;
  for (const run of base) {
    const runStart = cursor;
    const runEnd = cursor + run.text.length;
    cursor = runEnd;
    // Run entirely outside the selection — keep as-is.
    if (runEnd <= from || runStart >= to) {
      out.push(run);
      continue;
    }
    const localFrom = Math.max(from, runStart) - runStart;
    const localTo = Math.min(to, runEnd) - runStart;
    const before = run.text.slice(0, localFrom);
    const covered = run.text.slice(localFrom, localTo);
    const after = run.text.slice(localTo);
    if (before) {
      const s = cloneRunStyle(run.style);
      out.push(s ? { text: before, style: s } : { text: before });
    }
    if (covered) {
      const transformed = transformedSlices[transformedIndex] ?? "";
      transformedIndex += 1;
      // A transform may delete all characters of a run only if it returned ""
      // (it never does for these case maps), but guard anyway — drop an empty.
      if (transformed) {
        const s = cloneRunStyle(run.style);
        out.push(s ? { text: transformed, style: s } : { text: transformed });
      }
    }
    if (after) {
      const s = cloneRunStyle(run.style);
      out.push(s ? { text: after, style: s } : { text: after });
    }
  }
  return mergeRuns(out);
}

/**
 * Convenience for the WHOLE-block case (the common path: no sub-range
 * selection): transform every character of `text`, returning the new flat
 * string. Equivalent to `transformCase(text, mode)` but named so the call site
 * reads symmetrically with {@link applyCaseToRuns}. PURE.
 */
export function applyCaseToText(text: string, mode: ChangeCaseMode): string {
  return transformCase(text, mode);
}

// ===========================================================================
// PASTE TEXT FROM THE OS CLIPBOARD — pure sanitizer + gate
// ===========================================================================

/**
 * C0/C1 control bytes to strip from pasted clipboard text, EXCEPT tab (U+0009)
 * and newline (U+000A) which are kept so a multi-line paste survives. Carriage
 * returns are normalized to `\n` before this runs (so U+000D is in the set).
 * Mirrors `presentation-text-format.SHAPE_TEXT_CONTROL_RE`.
 */
const CLIPBOARD_TEXT_CONTROL_RE = /[\u0000-\u0008\u000b-\u001f\u007f-\u009f]/g;

/**
 * The length cap for a single pasted text block. Generous (a long pasted
 * paragraph / outline is fine) but bounded so a hostile multi-megabyte
 * clipboard can't land one giant block on the slide. The persisted runs are
 * additionally capped per-run by `presentation-runs` at the serialize edge.
 */
export const MAX_PASTED_TEXT_LENGTH = 100_000;

/**
 * Sanitize arbitrary OS-clipboard text into the string to seed a new text box
 * with, or `null` when it carries nothing usable (so the caller does NOT create
 * an empty box). Normalizes CRLF → LF, strips control bytes (keeping tab /
 * newline), and caps the length. This is the GATE that lets a Ctrl/Cmd+V with
 * no in-app block payload and no image fall through to a text paste only when
 * there is real text. PURE; never throws, never mutates.
 */
export function sanitizeClipboardText(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const cleaned = value
    .replace(/\r\n?/g, "\n")
    .replace(CLIPBOARD_TEXT_CONTROL_RE, "")
    .slice(0, MAX_PASTED_TEXT_LENGTH);
  // Keep the interior whitespace/newlines, but a value that is ONLY whitespace
  // carries no content → no box.
  return cleaned.trim() === "" ? null : cleaned;
}

// ===========================================================================
// LOCK OBJECT — PowerPoint's Lock: the frame stays put, the content does not
// ===========================================================================
//
// PowerPoint's Lock (right-click > Lock, the Selection Pane's padlock) fixes
// an object's place: it can no longer be moved, sized or turned. Microsoft
// Support: a locked object's position is fixed and its content and format
// can still be changed; it can still be deleted (Office-Watch; Microsoft's
// page is silent). What Lock writes says the rest: `a:spLocks noGrp noRot
// noMove noResize noEditPoints noAdjustHandles noChangeArrowheads
// noChangeShapeType` (a group's `noGrp noUngrp noRot noMove noResize`), so a
// locked object also does not join or leave a group, change its shape, edit
// its points, drag its adjustment handles or change its arrowheads -- and
// nothing else: typing, Format Shape, picture styles, text effects, Delete,
// Cut, the Selection Pane's name and eye, and the stacking order stay
// open. Word's and Excel's locks are other rules (a Word anchor lock, an
// Excel object's Locked under sheet protection) and do not come here.

/** A minimal positioned-block shape that may carry the additive `locked` flag. */
export interface LockableBlock {
  locked?: unknown;
  [key: string]: unknown;
}

/**
 * The single source of truth for "is this block locked?" (the editor's
 * Lock / Unlock). A block is locked only when its `locked` field is STRICTLY
 * `true` (so a stray truthy value / absent field is treated as unlocked —
 * additive + migration-safe). PURE.
 */
export function isBlockLocked(
  block: LockableBlock | null | undefined,
): boolean {
  return !!block && block.locked === true;
}

/**
 * What a lock can refuse on an object: its frame (move, resize, rotate), its
 * place in a group (group, ungroup), its outline (change shape, edit points,
 * adjustment handles, arrowheads), its picture's crop and its text. Each is
 * one `a:spLocks`-family attribute (ECMA-376 20.1.2.2.34 and kin).
 */
export type PresentationLockOperation =
  | "move"
  | "resize"
  | "rotate"
  | "group"
  | "ungroup"
  | "changeShape"
  | "editPoints"
  | "adjust"
  | "arrowheads"
  | "crop"
  | "textEdit";

const LOCK_ATTRIBUTE: Record<PresentationLockOperation, string> = {
  move: "noMove",
  resize: "noResize",
  rotate: "noRot",
  group: "noGrp",
  ungroup: "noUngrp",
  changeShape: "noChangeShapeType",
  editPoints: "noEditPoints",
  adjust: "noAdjustHandles",
  arrowheads: "noChangeArrowheads",
  crop: "noCrop",
  textEdit: "noTextEdit",
};

/**
 * What the editor's Lock refuses: what PowerPoint's Lock writes (above). A
 * crop is a new frame for the picture (its `a:xfrm` moves with the crop), so
 * a locked picture keeps its crop too. Text is never Lock's.
 */
const LOCK_REFUSES: ReadonlySet<PresentationLockOperation> = new Set([
  "move",
  "resize",
  "rotate",
  "group",
  "ungroup",
  "changeShape",
  "editPoints",
  "adjust",
  "arrowheads",
  "crop",
]);

/** What Unlock takes off a file's locks (what Lock writes); an aspect lock,
 * a picture's `noCrop` and Designer's `noTextEdit` stay. */
const UNLOCK_CLEARS: ReadonlySet<PresentationLockOperation> = new Set([
  "move",
  "resize",
  "rotate",
  "group",
  "ungroup",
  "changeShape",
  "editPoints",
  "adjust",
  "arrowheads",
]);

/**
 * Whether `operation` on `block` is refused by a lock: the editor's Lock
 * (`locked: true`) or the file's own frame locks (`a:spLocks` and its
 * picture, connector, group and frame kin, read as `source.nativeFrameLocks`;
 * the group's own as `source.nativeGroupFrameLocks`, which `wholeGroup`
 * gestures read). PowerPoint neither moves, sizes nor turns such an object
 * (Designer's design elements lock all three) and still formats or deletes
 * it. Unlock on an object its file locks (`locked: false`) takes off what
 * Lock writes, and keeps the file's crop and text locks. PURE.
 */
export function isLockedFor(
  block: { locked?: unknown; source?: unknown } | null | undefined,
  operation: PresentationLockOperation,
  wholeGroup = false,
): boolean {
  if (!block) return false;
  if (block.locked === true && LOCK_REFUSES.has(operation)) return true;
  if (operation === "crop" && isLockedFor(block, "resize", wholeGroup)) return true;
  if (block.locked === false && UNLOCK_CLEARS.has(operation)) return false;
  const source = block.source;
  if (!source || typeof source !== "object" || Array.isArray(source)) return false;
  const locks = (source as Record<string, unknown>)[wholeGroup ? "nativeGroupFrameLocks" : "nativeFrameLocks"];
  return Array.isArray(locks) && locks.includes(LOCK_ATTRIBUTE[operation]);
}

/**
 * Whether the file locks `operation` on `block`'s frame (see
 * {@link isLockedFor}; the editor's own Lock is read by
 * {@link canManipulateBlock}). PURE.
 */
export function isFrameLockedByFile(
  block: { source?: unknown } | null | undefined,
  operation: "move" | "resize" | "rotate",
  wholeGroup = false,
): boolean {
  if (!block) return false;
  const { locked } = block as { locked?: unknown };
  return isLockedFor({ locked: locked === false ? false : undefined, source: block.source }, operation, wholeGroup);
}

/** Which lock each patched field of a block answers to. */
const PATCH_LOCKS: Readonly<Record<string, PresentationLockOperation>> = {
  left: "move",
  top: "move",
  width: "resize",
  height: "resize",
  rotation: "rotate",
  flipH: "rotate",
  flipV: "rotate",
  cropRect: "crop",
  headEnd: "arrowheads",
  tailEnd: "arrowheads",
  headEndWidth: "arrowheads",
  headEndLength: "arrowheads",
  tailEndWidth: "arrowheads",
  tailEndLength: "arrowheads",
  shape: "changeShape",
  adjustments: "adjust",
  customPath: "editPoints",
  points: "editPoints",
};

/**
 * Whether a lock refuses `patch` on `block`: a field the patch changes is
 * the frame's (place, size, turn, flip, crop) or the outline's (shape,
 * adjustments, points, arrowheads), and that is locked ({@link isLockedFor}). Fill,
 * outline, effects, text and the rest are never refused by a lock. PURE.
 */
export function lockRefusesPatch(
  block: { locked?: unknown; source?: unknown } | null | undefined,
  patch: Readonly<Record<string, unknown>>,
): boolean {
  if (!block) return false;
  const current = block as Record<string, unknown>;
  return Object.entries(patch).some(([key, value]) => {
    const operation = PATCH_LOCKS[key];
    return operation !== undefined && JSON.stringify(value) !== JSON.stringify(current[key]) && isLockedFor(block, operation);
  });
}

/**
 * Whether the object's FRAME may change: drag-to-move, handle-resize,
 * rotate, flip, nudge, align / distribute, Size & Position, crop, group /
 * ungroup, Change Shape, Edit Points, adjustment handles. A missing block
 * can't; a locked one (Lock) is fixed in place. Content and format edits and
 * Delete do NOT read this -- they read {@link canEditBlockContent}. PURE.
 */
export function canManipulateBlock(
  block: LockableBlock | null | undefined,
): boolean {
  return !!block && !isBlockLocked(block);
}

/**
 * Whether `locked: true` is the import keeping an object read-only rather
 * than Lock: the page marks so what no save can write back -- a SmartArt
 * member kept as the file drew it (`nativeGroupDisposition: "preserve-only"`),
 * an add-in's fallback picture, a narration icon, and an imported object
 * that names no shape of its slide (a layout's diagram drawn on the slide, a
 * picture known only by its rendering). A model object's `locked` (no
 * native source) and an imported shape's (Lock, or the file's Lock read
 * back) are Lock. PURE.
 */
export function isBlockKeptReadOnly(block: { locked?: unknown; source?: unknown } | null | undefined): boolean {
  if (!block || block.locked !== true) return false;
  const record = block as Record<string, unknown>;
  if (record.nativeGroupDisposition === "preserve-only") return true;
  const source = block.source;
  if (!source || typeof source !== "object" || Array.isArray(source)) return false;
  const native = source as Record<string, unknown>;
  if (native.nativeAddInFallback !== undefined || native.nativeNarrationIcon === true) return true;
  if (isTextKeptReadOnly(block)) return true;
  return typeof native.shapeId !== "string" || native.shapeId === "";
}

/**
 * Whether the object is a text shape whose words the import could not make
 * editable and keeps read-only (PPTX reading row 814,
 * `source.nativeTextKeptReadOnly`): drawn from the file's runs, sizes and
 * frame, never written back -- its shape stays as the file has it. Before
 * the row such a shape was not on the page at all. PURE.
 */
export function isTextKeptReadOnly(block: { locked?: unknown; source?: unknown } | null | undefined): boolean {
  if (!block || block.locked !== true) return false;
  const source = block.source;
  if (!source || typeof source !== "object" || Array.isArray(source)) return false;
  const kept = (source as Record<string, unknown>).nativeTextKeptReadOnly;
  return !!kept && typeof kept === "object" && !Array.isArray(kept);
}

/**
 * Whether a read-only text shape's words were cut at the page's bound (PPTX
 * reading row 824, `source.nativeTextKeptReadOnly.truncated`): only their
 * first part is drawn, the file keeps them all. PURE.
 */
export function isTextKeptReadOnlyTruncated(block: { locked?: unknown; source?: unknown } | null | undefined): boolean {
  if (!isTextKeptReadOnly(block)) return false;
  const kept = (block!.source as Record<string, unknown>).nativeTextKeptReadOnly as Record<string, unknown>;
  return kept.truncated === true;
}

/**
 * Whether the object's CONTENT and FORMAT may change, or the object be
 * deleted or cut: text, fills, outlines, effects, picture styles and
 * adjustments, table cells, chart data, SmartArt text and styles, alt text,
 * links, animation, the Selection Pane's name and visibility. PowerPoint's
 * Lock does not refuse any of these (a locked object's position is fixed and
 * its content and format can still be changed); only an object the import
 * keeps read-only ({@link isBlockKeptReadOnly}) refuses them. PURE.
 */
export function canEditBlockContent(
  block: LockableBlock | null | undefined,
): boolean {
  return !!block && !isBlockKeptReadOnly(block);
}

/**
 * Whether the object may change its place in the stacking order (Bring
 * Forward / Send Backward / to Front / to Back, the Selection Pane's arrows
 * and drag, Reorder Objects), and whether another object may step past it.
 * PowerPoint's Lock does not hold the order: `a:spLocks` has no stacking
 * attribute (ECMA-376 20.1.2.2.34), and a locked object "still can be ...
 * sent to back/front" (Microsoft Q&A, the Selection Pane's Lock). Only an
 * object the import keeps read-only ({@link isBlockKeptReadOnly}) keeps its
 * depth. PURE.
 */
export function canReorderBlock(
  block: LockableBlock | null | undefined,
): boolean {
  return canEditBlockContent(block);
}

/**
 * Whether typing into the object is refused: never by Lock; by the file's
 * `noTextEdit` (Designer's design elements), and on an object the import
 * keeps read-only. PURE.
 */
export function isTextEditLocked(block: { locked?: unknown; source?: unknown } | null | undefined): boolean {
  return isBlockKeptReadOnly(block) || isLockedFor(block, "textEdit");
}

/**
 * Whether the object reads as locked (Lock / Unlock): locked in the editor,
 * or by its file (any of its frame's move, size or turn locks; PowerPoint's
 * Lock writes them all). PURE.
 */
export function isBlockShownLocked(block: (LockableBlock & { source?: unknown }) | null | undefined): boolean {
  if (!block) return false;
  if (isBlockLocked(block)) return true;
  return (["move", "resize", "rotate"] as const).some((operation) => isFrameLockedByFile(block, operation));
}

/**
 * Toggle the `locked` flag for a block, returned as a fresh object (never
 * mutates the input). Locking SETS `locked: true`; unlocking DELETES the key so
 * an unlocked block stays byte-for-byte identical to one that was never locked
 * (no `locked: false` residue). PURE.
 */
export function withToggledLock<T extends LockableBlock & { source?: unknown }>(block: T): T {
  const next = { ...block } as T;
  if (isBlockShownLocked(block)) {
    // Unlocking an object its file locks says so (`locked: false`): the
    // save takes the file's locks off, and the editor stops honouring them.
    const byFile = (["move", "resize", "rotate"] as const).some((operation) =>
      isFrameLockedByFile({ ...block, locked: undefined } as T, operation),
    );
    if (byFile) (next as LockableBlock).locked = false;
    else delete (next as LockableBlock).locked;
  } else {
    (next as LockableBlock).locked = true;
  }
  return next;
}








export function toFullWidth(text: string): string {
  let out = "";
  for (const ch of text) {
    const code = ch.codePointAt(0) ?? 0;
    if (code === 0x20) out += "\u3000";
    else if (code >= 0x21 && code <= 0x7e)
      out += String.fromCodePoint(code + 0xfee0);
    else out += ch;
  }
  return out;
}

export function toHalfWidth(text: string): string {
  let out = "";
  for (const ch of text) {
    const code = ch.codePointAt(0) ?? 0;
    if (code === 0x3000) out += " ";
    else if (code >= 0xff01 && code <= 0xff5e)
      out += String.fromCodePoint(code - 0xfee0);
    else out += ch;
  }
  return out;
}
