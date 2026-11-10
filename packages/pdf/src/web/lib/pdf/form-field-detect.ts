/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { VectorMark } from "./vector-marks";
/**
 * Propose fillable fields for a FLAT PDF.
 *
 * A worksheet exported to PDF, or a scan run through OCR, has no AcroForm at
 * all: the blanks are underscores and the tick boxes are glyphs. Acrobat's
 * Prepare Form and WPS's form tool both convert those into real widgets, and
 * without it the only way to fill such a page here is to draw text boxes by
 * hand over every blank.
 *
 * This is the detection half: a pure pass over the extracted text boxes that
 * PROPOSES candidates with a reason and a confidence, for a human to review.
 * It deliberately does not create anything — a heuristic that silently writes
 * widgets onto a document produces a form nobody asked for, and the wrong
 * guesses are invisible until someone tries to fill it.
 *
 * Geometry is the extractor's own: page-relative fractions in [0, 1] with a
 * top-left origin, so a candidate can be handed straight back to the overlay.
 */

/** A box as the text extractor produces it — only the fields detection reads. */
export interface DetectionTextBox {
  /** 1-based page number. */
  page: number;
  /** Page-relative fractions (0..1), top-left origin. */
  x: number;
  y: number;
  width: number;
  height: number;
  text: string;
}

export type FieldCandidateReason =
  | "underscore-run"
  | "label-colon"
  | "checkbox-glyph"
  | "ruled-line"
  | "drawn-box";

export interface FieldCandidate {
  kind: "text" | "checkbox";
  page: number;
  x: number;
  y: number;
  width: number;
  height: number;
  /** Suggested field name, unique within the returned set. */
  name: string;
  /** The label text it was derived from, shown in review. */
  label: string;
  /** Why it was proposed — a reviewer can judge a rule they can see. */
  reason: FieldCandidateReason;
  /** 0..1. Higher is a stronger signal, not a probability. */
  confidence: number;
}

export interface DetectFormFieldsOptions {
  /**
   * Vector rules and boxes recovered from the page's operator list. Real
   * forms draw their blanks and tick boxes; none of that reaches the text
   * layer, which is why the text-only detector misses exactly the forms
   * people most want filled.
   */
  vectorMarks?: ReadonlyArray<VectorMark & { page: number }>;
  /** Cap on proposals, so a pathological page cannot produce thousands. */
  maxCandidates?: number;
  /** Minimum run of underscores that counts as a blank. */
  minUnderscoreRun?: number;
  /** Width given to a label-derived field when nothing bounds it, as a
   *  fraction of the page. */
  defaultFieldWidth?: number;
}

const DEFAULTS = {
  maxCandidates: 200,
  minUnderscoreRun: 3,
  defaultFieldWidth: 0.22,
};

/** Glyphs a flat form uses for an empty tick box. */
const CHECKBOX_GLYPHS = ["□", "☐", "▢", "◻", "○", "◯"];
/** Bracket pairs that read as a tick box when they enclose only spaces. */
const CHECKBOX_BRACKETS = /\[\s*\]|\(\s*\)|（\s*）|【\s*】/g;

/** Trailing punctuation that marks a label rather than prose. */
const LABEL_SUFFIX = /[:：]\s*$/;

/**
 * Turn a label into a field name: keep letters, digits and CJK, collapse the
 * rest to underscores. The name is what the filled value persists under, so it
 * has to be stable and free of the whitespace an AcroForm name may legally
 * contain but which makes a value map painful to read.
 */
export function fieldNameFromLabel(label: string, fallback: string): string {
  const cleaned = label
    .replace(LABEL_SUFFIX, "")
    .replace(/[_\s]+/g, " ")
    .trim()
    .replace(/[^\p{L}\p{N}]+/gu, "_")
    .replace(/^_+|_+$/g, "");
  return cleaned.slice(0, 60) || fallback;
}

function uniqueName(base: string, taken: Set<string>): string {
  if (!taken.has(base)) {
    taken.add(base);
    return base;
  }
  for (let i = 2; i < 1000; i += 1) {
    const candidate = `${base}_${i}`;
    if (!taken.has(candidate)) {
      taken.add(candidate);
      return candidate;
    }
  }
  return base;
}

/** Two boxes sit on the same line when their vertical spans mostly overlap. */
function sameLine(a: DetectionTextBox, b: DetectionTextBox): boolean {
  const top = Math.max(a.y, b.y);
  const bottom = Math.min(a.y + a.height, b.y + b.height);
  const overlap = bottom - top;
  return overlap > Math.min(a.height, b.height) * 0.5;
}

/** Horizontal slice of a box covering characters [from, to) of its text. */
function sliceRect(
  box: DetectionTextBox,
  from: number,
  to: number,
): { x: number; width: number } | null {
  const length = box.text.length;
  if (length === 0 || to <= from) return null;
  const unit = box.width / length;
  return { x: box.x + unit * from, width: unit * (to - from) };
}

/**
 * Propose fillable fields over a page's extracted text.
 *
 * Three rules, each independently visible in the result's `reason`:
 *
 *  - a run of underscores is a blank to type on;
 *  - a label ending in ':' with clear space to its right is a blank beside it;
 *  - an empty tick-box glyph or bracket pair is a checkbox.
 *
 * Candidates are returned strongest-first and never overlap: a line that is
 * both "Name:" and "____" yields one field, not two stacked on each other.
 */
export function detectFormFields(
  boxes: ReadonlyArray<DetectionTextBox>,
  options: DetectFormFieldsOptions = {},
): FieldCandidate[] {
  const config = { ...DEFAULTS, ...options };
  const usable = boxes.filter(
    (box) =>
      typeof box.text === "string" &&
      Number.isFinite(box.x) &&
      Number.isFinite(box.y) &&
      box.width > 0 &&
      box.height > 0,
  );
  const found: FieldCandidate[] = [];
  const taken = new Set<string>();

  for (const box of usable) {
    // --- checkboxes ------------------------------------------------------
    for (const glyph of CHECKBOX_GLYPHS) {
      let at = box.text.indexOf(glyph);
      while (at !== -1) {
        const slice = sliceRect(box, at, at + glyph.length);
        if (slice) {
          const size = Math.min(box.height, slice.width || box.height);
          found.push({
            kind: "checkbox",
            page: box.page,
            x: slice.x,
            y: box.y,
            width: size,
            height: size,
            name: "",
            label: labelAround(box.text, at, at + glyph.length),
            reason: "checkbox-glyph",
            confidence: 0.8,
          });
        }
        at = box.text.indexOf(glyph, at + glyph.length);
      }
    }
    CHECKBOX_BRACKETS.lastIndex = 0;
    let bracket: RegExpExecArray | null = CHECKBOX_BRACKETS.exec(box.text);
    while (bracket !== null) {
      const slice = sliceRect(box, bracket.index, bracket.index + bracket[0].length);
      if (slice) {
        const size = Math.min(box.height, slice.width || box.height);
        found.push({
          kind: "checkbox",
          page: box.page,
          x: slice.x,
          y: box.y,
          width: size,
          height: size,
          name: "",
          label: labelAround(
            box.text,
            bracket.index,
            bracket.index + bracket[0].length,
          ),
          reason: "checkbox-glyph",
          confidence: 0.7,
        });
      }
      bracket = CHECKBOX_BRACKETS.exec(box.text);
    }

    // --- underscore blanks -----------------------------------------------
    const runs = [...box.text.matchAll(/_{2,}/g)];
    for (const run of runs) {
      if (run[0].length < config.minUnderscoreRun) continue;
      const start = run.index ?? 0;
      const slice = sliceRect(box, start, start + run[0].length);
      if (!slice) continue;
      found.push({
        kind: "text",
        page: box.page,
        x: slice.x,
        y: box.y,
        width: slice.width,
        height: box.height,
        name: "",
        label: box.text.slice(0, start).trim(),
        reason: "underscore-run",
        // A long rule is a stronger signal than a short one.
        confidence: Math.min(0.95, 0.6 + run[0].length * 0.02),
      });
    }

    // --- "Label:" with space to the right ---------------------------------
    if (LABEL_SUFFIX.test(box.text) && runs.length === 0) {
      const rightEdge = box.x + box.width;
      const neighbour = usable
        .filter(
          (other) =>
            other !== box && other.page === box.page && sameLine(other, box) && other.x >= rightEdge,
        )
        .sort((a, b) => a.x - b.x)[0];
      const available = (neighbour ? neighbour.x : 1) - rightEdge;
      const width = Math.min(
        config.defaultFieldWidth,
        Math.max(0, available - 0.005),
      );
      // Under a couple of percent of the page there is nowhere to type; a label
      // immediately followed by its value is prose, not a blank.
      if (width > 0.02) {
        found.push({
          kind: "text",
          page: box.page,
          x: rightEdge + 0.005,
          y: box.y,
          width,
          height: box.height,
          name: "",
          label: box.text.trim(),
          reason: "label-colon",
          confidence: neighbour ? 0.55 : 0.65,
        });
      }
    }
  }

  // --- drawn rules and boxes ---------------------------------------------
  for (const mark of options.vectorMarks ?? []) {
    const lineOf = (other: DetectionTextBox) => {
      const top = Math.max(other.y, mark.y - 0.02);
      const bottom = Math.min(other.y + other.height, mark.y + mark.height + 0.005);
      return bottom - top > 0;
    };
    if (mark.kind === "rule" || mark.height <= 0.004) {
      // The blank after "Name:" drawn as a line: the field sits ON the rule,
      // labelled by the nearest text ending to its left.
      const label = usable
        .filter(
          (other) =>
            other.page === mark.page &&
            lineOf(other) &&
            other.x + other.width <= mark.x + 0.02,
        )
        .sort(
          (a, b) => b.x + b.width - (a.x + a.width),
        )[0];
      const height = label?.height ?? 0.024;
      found.push({
        kind: "text",
        page: mark.page,
        x: mark.x,
        y: mark.y + mark.height - height,
        width: mark.width,
        height,
        name: "",
        label: label?.text.trim() ?? "",
        reason: "ruled-line",
        // A rule with a labelled left end is the classic form blank; a bare
        // rule could be a divider, and the reviewer sees both.
        confidence: label && LABEL_SUFFIX.test(label.text) ? 0.85 : label ? 0.7 : 0.5,
      });
      continue;
    }
    // A drawn tick box, labelled by the nearest text beside it.
    const beside = usable
      .filter((other) => other.page === mark.page && lineOf(other))
      .sort(
        (a, b) =>
          Math.abs(a.x - (mark.x + mark.width)) -
          Math.abs(b.x - (mark.x + mark.width)),
      )[0];
    found.push({
      kind: "checkbox",
      page: mark.page,
      x: mark.x,
      y: mark.y,
      width: mark.width,
      height: mark.height,
      name: "",
      label: beside?.text.trim().slice(0, 80) ?? "",
      reason: "drawn-box",
      confidence: beside ? 0.75 : 0.55,
    });
  }

  const ordered = found
    .sort((a, b) => b.confidence - a.confidence)
    .filter((candidate, index, all) =>
      all.every(
        (other, otherIndex) => otherIndex >= index || !overlaps(candidate, other),
      ),
    )
    .slice(0, config.maxCandidates);

  // Reading order for review; the confidence that decided the dedupe stays on
  // each record so the UI can still sort or pre-select by it.
  return ordered
    .sort((a, b) => a.page - b.page || a.y - b.y || a.x - b.x)
    .map((candidate) => ({
      ...candidate,
      name: uniqueName(
        fieldNameFromLabel(
          candidate.label,
          candidate.kind === "checkbox" ? "check" : "field",
        ),
        taken,
      ),
    }));
}

/** Whether two candidates cover enough of the same space to be duplicates. */
function overlaps(a: FieldCandidate, b: FieldCandidate): boolean {
  if (a.page !== b.page) return false;
  const x = Math.max(0, Math.min(a.x + a.width, b.x + b.width) - Math.max(a.x, b.x));
  const y = Math.max(0, Math.min(a.y + a.height, b.y + b.height) - Math.max(a.y, b.y));
  const intersection = x * y;
  if (intersection <= 0) return false;
  const smaller = Math.min(a.width * a.height, b.width * b.height);
  return smaller > 0 && intersection / smaller > 0.5;
}

/** Where the next tick box starts in `text` at or after `from`, or -1. */
function nextCheckboxAt(text: string, from: number): number {
  let best = -1;
  for (const glyph of CHECKBOX_GLYPHS) {
    const at = text.indexOf(glyph, from);
    if (at !== -1 && (best === -1 || at < best)) best = at;
  }
  const brackets = /\[\s*\]|\(\s*\)|（\s*）|【\s*】/g;
  brackets.lastIndex = from;
  const match = brackets.exec(text);
  if (match && (best === -1 || match.index < best)) best = match.index;
  return best;
}

/**
 * The words that label a tick box. A row of options — "□ Yes □ No □ N/A" — is
 * one text box, so the label has to stop at the NEXT box; taking everything to
 * the right names the first option after all three.
 */
function labelAround(text: string, from: number, to: number): string {
  const stop = nextCheckboxAt(text, to);
  const after = text.slice(to, stop === -1 ? undefined : stop).trim();
  if (after) return after.slice(0, 80);
  // Nothing to the right: the box is labelled from the left ("Agree □"), back
  // to the previous tick box if the line holds several.
  const previous = nextCheckboxAt(text.slice(0, from), 0);
  const beforeStart = previous === -1 ? 0 : previous;
  return text.slice(beforeStart, from).replace(/^[^\p{L}\p{N}]+/u, "").trim().slice(0, 80);
}
