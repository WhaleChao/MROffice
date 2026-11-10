/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Find & Replace over the editable PDF text-box model (WPS/Acrobat parity).
 *
 * The inline editor (`PdfTextEditor`) holds the page's editable text as a list
 * of boxes — each an extracted line or a user-created box — whose EFFECTIVE
 * text is `edit.text ?? box.text`. These pure helpers find every occurrence of
 * a query across those box texts (reusing the same `findMatches` engine the
 * find bar uses) and produce the updated `PdfTextBoxEdit[]` for replacing one
 * match or all matches.
 *
 * Everything here is framework-free (no React, no pdf.js, no DOM) so the
 * replace math is unit-testable in isolation; the component only owns the
 * editable boxes + the active-match index and renders the bar.
 */

import { findMatches, type PdfTextSearchOptions } from "@mrnote/core/web/lib/pdf/text-search.ts";
import { mathTextReplaced } from "@mrnote/core/web/lib/office-math-text-replace.ts";

import { latexForPdfFormulaTextEdit, preservePdfFormulaLatexForTextEdit } from "@mrnote/core/web/lib/pdf/formula-latex.ts";
import { effectiveTextBoxKind, hasFormulaLatex } from "@mrnote/core/web/lib/pdf/text-box-kind.ts";

/** The minimal box shape the replace engine reads (id + current text). */
export interface ReplaceBox {
  id: string;
  /** The EFFECTIVE text of the box (edit override already applied). */
  text: string;
  kind?: "text" | "formula";
  latex?: string;
}

/**
 * The minimal persisted-edit shape the replace engine reads / writes. The apply
 * functions are GENERIC over the concrete edit type `E extends ReplaceEdit`, so
 * a caller passing a full `PdfTextBoxEdit` (geometry, font, origin…) gets the
 * full object back with only `text` updated — every other field rides through
 * untouched via the object spread.
 */
export interface ReplaceEdit {
  id: string;
  page: number;
  text: string;
  kind?: "text" | "formula";
  latex?: string;
}

/** One occurrence of the query inside a specific box, in document order. */
export interface BoxTextMatch {
  /** Id of the box the match is in. */
  boxId: string;
  /** Inclusive start offset within that box's effective text. */
  start: number;
  /** Exclusive end offset within that box's effective text. */
  end: number;
}

/**
 * Find every occurrence of `query` across the ordered `boxes`, flattened into a
 * single document-ordered list (box order first, then ascending offset within a
 * box). Empty / whitespace-only queries and empty box lists yield `[]`.
 *
 * Reuses {@link findMatches} per box so the matching semantics (overlap,
 * whole-word, case-sensitivity) are identical to the read-only find bar.
 */
export function findBoxMatches(
  boxes: ReadonlyArray<ReplaceBox>,
  query: string,
  options: PdfTextSearchOptions = {},
): BoxTextMatch[] {
  if (!Array.isArray(boxes) || boxes.length === 0) return [];
  const out: BoxTextMatch[] = [];
  for (const box of boxes) {
    if (!box || typeof box.id !== "string") continue;
    const text = typeof box.text === "string" ? box.text : "";
    if (text.length === 0) continue;
    // findMatches takes an array of page texts; one box = one "page".
    const hits = findMatches([text], query, options);
    for (const hit of hits) {
      out.push({ boxId: box.id, start: hit.start, end: hit.end });
    }
  }
  return out;
}

/**
 * Replace a single half-open range `[start, end)` in `text` with `replacement`.
 * Out-of-range / inverted ranges return the text unchanged (defensive — the
 * caller derives ranges from {@link findBoxMatches} against the same text).
 */
export function replaceRange(
  text: string,
  start: number,
  end: number,
  replacement: string,
): string {
  if (
    typeof text !== "string" ||
    !Number.isInteger(start) ||
    !Number.isInteger(end) ||
    start < 0 ||
    end > text.length ||
    end < start
  ) {
    return typeof text === "string" ? text : "";
  }
  return text.slice(0, start) + replacement + text.slice(end);
}

/**
 * Replace EVERY non-overlapping occurrence of `query` in `text` with
 * `replacement`, honoring the same options as the find engine. Walks matches
 * right-to-left so earlier offsets stay valid as the string changes length.
 * Returns the original text when there are no matches (so callers can detect a
 * no-op and avoid creating a spurious edit).
 */
export function replaceAllInText(
  text: string,
  query: string,
  replacement: string,
  options: PdfTextSearchOptions = {},
): string {
  if (typeof text !== "string" || text.length === 0) return text;
  const matches = findMatches([text], query, options);
  if (matches.length === 0) return text;
  let next = text;
  for (let i = matches.length - 1; i >= 0; i -= 1) {
    next = replaceRange(next, matches[i].start, matches[i].end, replacement);
  }
  return next;
}

/**
 * Build the new edit for a box whose text changed, preserving every other
 * persisted field. When an edit already exists for the box it is cloned with
 * the new text; otherwise a minimal edit is created from `box`. Generic so the
 * concrete edit type (e.g. `PdfTextBoxEdit`) is preserved through the spread —
 * the caller's `editFor` is expected to supply a fully-seeded edit (geometry,
 * font, origin) for a box that has none yet, so the only branch that fabricates
 * a bare `{id,page,text}` is the defensive one where no seed was provided.
 */
function withReplacedText<E extends ReplaceEdit>(
  edit: E | undefined,
  box: ReplaceBox,
  page: number,
  newText: string,
  /** A formula's LaTeX already made to follow the text (`formulaRangeReplaced`). */
  followedLatex?: string,
): E {
  const next = edit
    ? { ...edit, text: newText }
    : ({ id: box.id, page, text: newText } as E);
  const kind = effectiveTextBoxKind(box, edit);
  if (kind !== "formula") {
    if (kind === "text" && next.latex !== undefined) {
      return { ...next, kind: "text", latex: undefined } as E;
    }
    return next as E;
  }
  return {
    ...next,
    kind: "formula",
    latex: followedLatex ?? latexForPdfFormulaTextEdit({
      currentLatex: edit?.latex ?? box.latex,
      previousText: edit?.text ?? box.text,
      nextText: newText,
    }),
  } as E;
}

/**
 * The LaTeX a formula box draws from, when it has one: a replacement in such
 * a box changes the LaTeX, or nothing (`formulaRangeReplaced`).
 */
function formulaLatexOf(box: ReplaceBox, edit: ReplaceEdit | undefined): string | null {
  if (effectiveTextBoxKind(box, edit) !== "formula") return null;
  const latex = edit?.latex ?? box.latex;
  return hasFormulaLatex(latex) ? (latex as string) : null;
}

/**
 * A formula's text with `[start, end)` replaced by `replacement`, its LaTeX
 * changed to match: the edit the formula editor makes of the same change
 * (`preservePdfFormulaLatexForTextEdit`), else the same characters changed in
 * the LaTeX (`mathTextReplaced`, the rule Word and PowerPoint Find & Replace
 * use). Null when the LaTeX cannot follow: the formula is then left exactly
 * as it is, never rebuilt from its plain text (a fraction would come back as
 * "a/b" characters).
 */
function formulaRangeReplaced(
  text: string,
  latex: string,
  start: number,
  end: number,
  replacement: string,
): { text: string; latex: string } | null {
  const next = replaceRange(text, start, end, replacement);
  if (next === text) return null;
  const preserved = preservePdfFormulaLatexForTextEdit({ currentLatex: latex, previousText: text, nextText: next });
  if (preserved !== null && preserved !== latex) return { text: next, latex: preserved };
  const edited = mathTextReplaced(text, latex, start, end, replacement);
  return edited ? { text: next, latex: edited.latex } : null;
}

/** A box (effective text) paired with the box's resolved page index. */
export interface ReplaceBoxWithPage extends ReplaceBox {
  page: number;
}

/** A find hit with the page its box is on. */
export interface PagedBoxTextMatch extends BoxTextMatch {
  page: number;
}






export function findDocumentBoxMatches(
  pages: ReadonlyArray<{ page: number; boxes: ReadonlyArray<ReplaceBox> }>,
  query: string,
  options: PdfTextSearchOptions = {},
): PagedBoxTextMatch[] {
  if (!query.trim()) return [];
  const out: PagedBoxTextMatch[] = [];
  for (const entry of [...pages].sort((a, b) => a.page - b.page)) {
    for (const match of findBoxMatches(entry.boxes, query, options)) {
      out.push({ ...match, page: entry.page });
    }
  }
  return out;
}

/**
 * The position of document hit `index` among the hits on its own page, or
 * -1 when there is no such hit. Replace-one acts on one page's boxes.
 */
export function pageLocalMatchIndex(
  matches: ReadonlyArray<PagedBoxTextMatch>,
  index: number,
): number {
  const match = matches[index];
  if (!match) return -1;
  let local = 0;
  for (let at = 0; at < index; at += 1) {
    if (matches[at].page === match.page) local += 1;
  }
  return local;
}

/**
 * Apply "replace one" to the match at flat index `matchIndex` (the active
 * find/replace match). Returns the updated edits list plus the resulting flat
 * matches AFTER the replace (so the caller can re-anchor the active index), or
 * `null` when the index is out of range / there is nothing to do.
 *
 * `editsById` maps a boxId → its existing edit (so the replacement preserves
 * geometry/font); `seedEdit` lets the component supply a freshly-seeded edit
 * (with origin/geometry) for a box that has none yet, so replacing inside an
 * unedited extracted line still masks the original glyphs on export.
 */
export function applyReplaceOne<E extends ReplaceEdit>(params: {
  boxes: ReadonlyArray<ReplaceBoxWithPage>;
  matchIndex: number;
  query: string;
  replacement: string;
  options?: PdfTextSearchOptions;
  editFor: (boxId: string) => E | undefined;
}): { editById: E; matchesAfter: BoxTextMatch[] } | { formulaLeft: true } | null {
  const { boxes, matchIndex, query, replacement, options, editFor } = params;
  const matches = findBoxMatches(boxes, query, options);
  if (matchIndex < 0 || matchIndex >= matches.length) return null;
  const target = matches[matchIndex];
  const box = boxes.find((b) => b.id === target.boxId);
  if (!box) return null;
  const current = editFor(target.boxId);
  const baseText =
    typeof current?.text === "string" ? current.text : box.text;
  const formulaLatex = formulaLatexOf(box, current);
  let newText: string;
  let followedLatex: string | undefined;
  if (formulaLatex !== null) {
    // Inside a formula: its LaTeX follows, or the formula is left whole.
    const edited = formulaRangeReplaced(baseText, formulaLatex, target.start, target.end, replacement);
    if (!edited) return { formulaLeft: true };
    newText = edited.text;
    followedLatex = edited.latex;
  } else {
    newText = replaceRange(baseText, target.start, target.end, replacement);
  }
  const editById = withReplacedText(current, box, box.page, newText, followedLatex);
  // Recompute matches against the post-replace box text so the caller can keep
  // the active index pointing at the next still-existing occurrence.
  const updatedBoxes = boxes.map((b) =>
    b.id === box.id ? { ...b, text: newText } : b,
  );
  const matchesAfter = findBoxMatches(updatedBoxes, query, options);
  return { editById, matchesAfter };
}

/**
 * Apply "replace all": replace every occurrence across all boxes and return the
 * list of changed edits (one per box that actually changed) plus the count of
 * matches replaced. Boxes with no match are untouched (not returned), so the
 * caller adds/overwrites only the changed edits.
 */
export function applyReplaceAll<E extends ReplaceEdit>(params: {
  boxes: ReadonlyArray<ReplaceBoxWithPage>;
  query: string;
  replacement: string;
  options?: PdfTextSearchOptions;
  editFor: (boxId: string) => E | undefined;
}): { changedEdits: E[]; replaced: number; formulaSkipped: number } {
  const { boxes, query, replacement, options, editFor } = params;
  // A box already edited (saved or not) is matched on the text it shows,
  // not the source extraction: counting on the source and bailing out when
  // it had no hit reported "no matches" for a word only the edit contains,
  // while the find bar and single Replace, which read the merged text, found
  
  const current = new Map<string, E | undefined>();
  const effective: ReplaceBoxWithPage[] = boxes.map((box) => {
    const edit = editFor(box.id);
    current.set(box.id, edit);
    return typeof edit?.text === "string" && edit.text !== box.text
      ? { ...box, text: edit.text }
      : box;
  });
  const matches = findBoxMatches(effective, query, options);
  if (matches.length === 0) return { changedEdits: [], replaced: 0, formulaSkipped: 0 };
  const changedEdits: E[] = [];
  // Matches inside a formula its LaTeX cannot follow: left, the formula whole.
  let formulaSkipped = 0;
  for (const box of effective) {
    const baseText = box.text;
    const formulaLatex = formulaLatexOf(box, current.get(box.id));
    if (formulaLatex !== null) {
      // Each match in turn, the last first so earlier offsets still hold.
      const inBox = findBoxMatches([box], query, options);
      let text = baseText;
      let latex = formulaLatex;
      for (let index = inBox.length - 1; index >= 0; index -= 1) {
        const edited = formulaRangeReplaced(text, latex, inBox[index].start, inBox[index].end, replacement);
        if (!edited) {
          formulaSkipped += 1;
          continue;
        }
        text = edited.text;
        latex = edited.latex;
      }
      if (text !== baseText) {
        changedEdits.push(withReplacedText(current.get(box.id), box, box.page, text, latex));
      }
      continue;
    }
    const newText = replaceAllInText(baseText, query, replacement, options);
    if (newText === baseText) continue; // no match in this box
    changedEdits.push(
      withReplacedText(current.get(box.id), box, box.page, newText),
    );
  }
  return { changedEdits, replaced: matches.length - formulaSkipped, formulaSkipped };
}
