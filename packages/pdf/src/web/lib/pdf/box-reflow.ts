/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Reflowing an edit through a chain of boxes (reflow R2).
 *
 * The chain is one text stream that happens to be stored in several boxes.
 * Editing a box re-fills that box and every box AFTER it: each keeps its own
 * geometry and font size, takes the lines that fit, and hands the remainder
 * on. What does not fit at the end is reported, never dropped — the caller
 * falls back to shrink-to-fit, which is today's behaviour.
 *
 * Two invariants the tests hold:
 *   - boxes BEFORE the edited one are untouched, byte for byte;
 *   - the concatenation of the outputs plus the overflow equals the
 *     concatenation of the inputs (with the edit applied) — reflow moves
 *     text, it never invents or loses any.
 */

export interface ReflowBox {
  id: string;
  /** Page fractions. */
  width: number;
  height: number;
  /** Fraction of page height. */
  fontSize: number;
  text: string;
}

/** Measures a string's width at a font size, both in page-fraction units. */
export type ReflowMeasure = (text: string, fontSize: number) => number;

export interface ReflowResult {
  /** New text per box id — only boxes that changed appear. */
  boxTexts: Map<string, string>;
  /** Text that fit nowhere; empty when the chain absorbed everything. */
  overflow: string;
  changedBoxIds: string[];
}

/** Line height as a multiple of font size — matches the export's ratio. */
const LINE_HEIGHT_RATIO = 1.25;

/**
 * Join box texts into one stream, healing hyphenation at the seams.
 *
 * Extraction stores a line-break hyphen as a literal "-" at the box end;
 * joining with a space would leave "contin- uation" in the reflowed output.
 * Only a lowercase continuation heals — "state-of-the-art" keeps its hyphen
 * when the next box starts a new word with a capital or a digit.
 */
export function joinBoxTexts(texts: ReadonlyArray<string>): string {
  let joined = "";
  for (const raw of texts) {
    const text = raw.trim();
    if (!text) continue;
    if (!joined) {
      joined = text;
      continue;
    }
    if (/[-‐]$/.test(joined) && /^[a-zà-öø-ÿ]/.test(text)) {
      joined = joined.replace(/[-‐]$/, "") + text;
    } else {
      joined = joinSegments(joined, text);
    }
  }
  return joined.replace(/\s+/g, " ").trim();
}

/** Split into wrap units: a Latin word, one CJK character, or one symbol. */
function wrapUnits(text: string): string[] {
  const units: string[] = [];
  const pattern = /[A-Za-zÀ-ÖØ-öø-ÿ0-9'’\-‐]+|[^\sA-Za-zÀ-ÖØ-öø-ÿ0-9]|\S/gu;
  for (const match of text.matchAll(pattern)) units.push(match[0]);
  return units;
}

/** Whether a unit is CJK-ish (joins without a space). */
function joinsWithoutSpace(unit: string): boolean {
  return !/^[A-Za-zÀ-ÖØ-öø-ÿ0-9'’\-‐]+$/.test(unit);
}

/**
 * Join two text segments the way the writing joins: no space when either side
 * of the seam is CJK. A space inside running CJK text is not "extra
 * whitespace", it is wrong typography that would land IN the box text.
 */
function joinSegments(a: string, b: string): string {
  if (!a) return b;
  if (!b) return a;
  const left = a[a.length - 1];
  const right = b[0];
  if (joinsWithoutSpace(left) || joinsWithoutSpace(right)) return a + b;
  return `${a} ${b}`;
}

function appendUnit(line: string, unit: string): string {
  if (!line) return unit;
  if (joinsWithoutSpace(unit) || line.endsWith("‐")) return line + unit;
  const last = line[line.length - 1];
  if (joinsWithoutSpace(last)) return line + unit;
  return `${line} ${unit}`;
}

/**
 * Fill one box: greedy word wrap against the box's own width, stopping at its
 * line capacity. Returns the text that fit and the remainder.
 */
export function fillBox(
  box: Pick<ReflowBox, "width" | "height" | "fontSize">,
  units: ReadonlyArray<string>,
  measure: ReflowMeasure,
): { fitted: string; rest: string[] } {
  const lineHeight = box.fontSize * LINE_HEIGHT_RATIO;
  const capacity = Math.max(1, Math.floor(box.height / lineHeight + 1e-6));
  const lines: string[] = [];
  let line = "";
  let index = 0;
  while (index < units.length) {
    const candidate = appendUnit(line, units[index]);
    if (measure(candidate, box.fontSize) <= box.width + 1e-9 || !line) {
      line = candidate;
      index += 1;
      continue;
    }
    lines.push(line);
    line = "";
    if (lines.length >= capacity) break;
  }
  if (line && lines.length < capacity) {
    lines.push(line);
    line = "";
  }
  const rest = line ? [line, ...units.slice(index)] : [...units.slice(index)];
  return {
    fitted: lines.reduce((acc, next) => joinSegments(acc, next), ""),
    rest,
  };
}

/**
 * Reflow `chain` after replacing `editedBoxId`'s text with `newText`.
 *
 * Boxes before the edited one are not in the result at all — untouched is the
 * strongest statement of unchanged.
 */
export function reflowChain(
  chain: ReadonlyArray<ReflowBox>,
  editedBoxId: string,
  newText: string,
  measure: ReflowMeasure,
): ReflowResult | null {
  const start = chain.findIndex((box) => box.id === editedBoxId);
  if (start < 0) return null;

  const downstream = chain.slice(start);
  const stream = joinBoxTexts([
    newText,
    ...downstream.slice(1).map((box) => box.text),
  ]);
  let units = wrapUnits(stream);

  const boxTexts = new Map<string, string>();
  const changedBoxIds: string[] = [];
  for (const box of downstream) {
    const { fitted, rest } = fillBox(box, units, measure);
    units = rest.flatMap((unit) => wrapUnits(unit));
    if (fitted !== box.text.trim() || box.id === editedBoxId) {
      boxTexts.set(box.id, fitted);
      changedBoxIds.push(box.id);
    }
    if (units.length === 0) {
      // Downstream boxes past the end of the stream empty out — their text
      // moved up. Recording them keeps the conservation invariant visible.
      continue;
    }
  }

  return {
    boxTexts,
    overflow: units.length
      ? units.reduce((acc, u) => joinSegments(acc, u), "")
      : "",
    changedBoxIds,
  };
}

/**
 * Fold a reflow's overflow into its LAST box.
 *
 * Applying a reflow and discarding the overflow deletes text — the exact
 * thing the conservation invariant forbids. The last box absorbs what fit
 * nowhere; it will overflow its rectangle, and the existing shrink-to-fit
 * handles an overfull box exactly as it does today without reflow.
 */
export function absorbReflowOverflow(
  result: ReflowResult,
  chain: ReadonlyArray<Pick<ReflowBox, "id" | "text">>,
): ReflowResult {
  if (!result.overflow || chain.length === 0) return result;
  const last = chain[chain.length - 1];
  const current = result.boxTexts.get(last.id) ?? last.text;
  const boxTexts = new Map(result.boxTexts);
  boxTexts.set(last.id, joinBoxTexts([current, result.overflow]));
  return {
    boxTexts,
    overflow: "",
    changedBoxIds: result.changedBoxIds.includes(last.id)
      ? result.changedBoxIds
      : [...result.changedBoxIds, last.id],
  };
}
