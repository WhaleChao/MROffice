/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Reflow that carries the styling with the text.
 *
 * The plain reflow moves characters between boxes and drops their runs — the
 * honest v1, since stale runs are worse than none. This carries them: every
 * non-whitespace character keeps the style it entered with, wherever it lands.
 *
 * The trick that keeps this simple is that reflow CONSERVES non-whitespace
 * characters in order — that is its tested invariant — while whitespace is
 * the wrap's own artefact. So styling is mapped per non-whitespace character:
 * build the style sequence of the input stream once, then walk each output
 * box and hand styles back out in order. No unit surgery, no offset algebra
 * through the wrap; whitespace takes the style of the character after it,
 * which is what a space between two words visually belongs to when they
 * differ.
 */

import {
  absorbReflowOverflow,
  reflowChain,
  type ReflowBox,
  type ReflowMeasure,
} from "./box-reflow";
import type { PdfTextStyleRun } from "@mrnote/core/web/lib/pdf/text-style-runs.ts";

/** The style of one character — a run minus its text. */
export type CharStyle = Omit<PdfTextStyleRun, "text">;

export interface StyledReflowBox extends ReflowBox {
  /** Valid runs for `text` (concatenation equals it), or absent. */
  runs?: PdfTextStyleRun[];
}

export interface StyledBoxOutput {
  text: string;
  /** Rebuilt runs; absent when the whole box is unstyled. */
  runs?: PdfTextStyleRun[];
}

export interface StyledReflowResult {
  outputs: Map<string, StyledBoxOutput>;
  changedBoxIds: string[];
}

/** Sentinel for "no styling recorded"; boxes made only of it drop runs. */
const UNSTYLED: CharStyle = {};

function styleKey(style: CharStyle): string {
  const entries = Object.entries(style)
    .filter(([, value]) => value !== undefined)
    .sort(([a], [b]) => (a < b ? -1 : 1));
  return JSON.stringify(entries);
}

/** Per-character styles for a box, from its runs or the sentinel. */
function charStyles(text: string, runs: PdfTextStyleRun[] | undefined): CharStyle[] {
  const chars = Array.from(text);
  if (!runs || runs.length === 0) {
    return chars.map(() => UNSTYLED);
  }
  const concatenated = runs.map((run) => run.text).join("");
  if (concatenated !== text) {
    // Invalid pair: treating it as unstyled is the same judgement the
    // persistence guards make, applied at the same boundary.
    return chars.map(() => UNSTYLED);
  }
  const out: CharStyle[] = [];
  for (const run of runs) {
    const { text: _text, ...style } = run;
    for (const _char of Array.from(run.text)) out.push(style);
  }
  return out;
}

/**
 * The style sequence of the joined stream's NON-WHITESPACE characters,
 * replicating the join's only character-dropping rule: a trailing hyphen
 * heals away before a lowercase continuation.
 */
export function buildStreamStyles(
  boxes: ReadonlyArray<{ text: string; runs?: PdfTextStyleRun[] }>,
): CharStyle[] {
  const out: CharStyle[] = [];
  const trimmed = boxes
    .map((box) => {
      // Offsets in CODE POINTS throughout: charStyles indexes by code point,
      // and mixing UTF-16 lengths in would misalign every style after an
      // astral character.
      const all = Array.from(box.text).length;
      const lead = all - Array.from(box.text.trimStart()).length;
      const kept = Array.from(box.text.trim()).length;
      return {
        text: box.text.trim(),
        styles: charStyles(box.text, box.runs).slice(lead, lead + kept),
      };
    })
    .filter((box) => box.text.length > 0);
  trimmed.forEach((box, index) => {
    let chars = Array.from(box.text);
    let styles = box.styles;
    const next = trimmed[index + 1];
    if (
      next &&
      /[-‐]$/.test(box.text) &&
      /^[a-zà-öø-ÿ]/.test(next.text)
    ) {
      // The healed hyphen never reaches the stream, so its style must not
      // either — otherwise every style after it is off by one character.
      chars = chars.slice(0, -1);
      styles = styles.slice(0, -1);
    }
    chars.forEach((char, at) => {
      if (!/\s/.test(char)) out.push(styles[at] ?? UNSTYLED);
    });
  });
  return out;
}

/** Rebuild runs for an output text by consuming stream styles in order. */
function rebuildRuns(
  text: string,
  take: () => CharStyle,
): PdfTextStyleRun[] | undefined {
  const chars = Array.from(text);
  const styles: CharStyle[] = new Array(chars.length);
  // Non-whitespace characters consume the stream; whitespace is resolved in a
  // second pass to the style of the character AFTER it.
  chars.forEach((char, at) => {
    if (!/\s/.test(char)) styles[at] = take();
  });
  for (let at = chars.length - 1; at >= 0; at -= 1) {
    if (styles[at] === undefined) {
      styles[at] = styles[at + 1] ?? styles[at - 1] ?? UNSTYLED;
    }
  }
  for (let at = 0; at < chars.length; at += 1) {
    if (styles[at] === undefined) styles[at] = UNSTYLED;
  }
  if (styles.every((style) => styleKey(style) === styleKey(UNSTYLED))) {
    return undefined;
  }
  const runs: PdfTextStyleRun[] = [];
  let current = "";
  let currentKey: string | null = null;
  let currentStyle: CharStyle = UNSTYLED;
  chars.forEach((char, at) => {
    const key = styleKey(styles[at]);
    if (currentKey === null || key === currentKey) {
      current += char;
      currentKey = key;
      currentStyle = styles[at];
      return;
    }
    runs.push({ text: current, ...currentStyle });
    current = char;
    currentKey = key;
    currentStyle = styles[at];
  });
  if (current) runs.push({ text: current, ...currentStyle });
  return runs;
}

/**
 * Reflow with styling carried along.
 *
 * Same contract as the plain version — untouched upstream, overflow absorbed
 * into the last box — plus: the i-th non-whitespace character of the output
 * carries the style of the i-th non-whitespace character of the input.
 */
export function reflowChainStyled(
  chain: ReadonlyArray<StyledReflowBox>,
  editedBoxId: string,
  newText: string,
  newRuns: PdfTextStyleRun[] | undefined,
  measure: ReflowMeasure,
): StyledReflowResult | null {
  const start = chain.findIndex((box) => box.id === editedBoxId);
  if (start < 0) return null;

  const raw = reflowChain(chain, editedBoxId, newText, measure);
  if (!raw) return null;
  const result = absorbReflowOverflow(raw, chain.slice(start));

  const streamStyles = buildStreamStyles([
    { text: newText, runs: newRuns },
    ...chain.slice(start + 1).map((box) => ({ text: box.text, runs: box.runs })),
  ]);
  let cursor = 0;
  const take = (): CharStyle => streamStyles[cursor++] ?? UNSTYLED;

  const outputs = new Map<string, StyledBoxOutput>();
  const changedBoxIds: string[] = [];
  for (const box of chain.slice(start)) {
    const text = result.boxTexts.get(box.id);
    if (text === undefined) {
      // Unchanged box: its characters still consume their styles, or every
      // box after it would wear its neighbour's formatting.
      for (const char of Array.from(box.text.trim())) {
        if (!/\s/.test(char)) cursor += 1;
      }
      continue;
    }
    outputs.set(box.id, { text, runs: rebuildRuns(text, take) });
    changedBoxIds.push(box.id);
  }
  return { outputs, changedBoxIds };
}
