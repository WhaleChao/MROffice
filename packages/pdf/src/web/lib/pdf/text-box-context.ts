/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { TextBoxEditLike } from "./text-box-flatten";
import { isFormulaTextBoxKind } from "@mrnote/core/web/lib/pdf/text-box-kind.ts";
import {
  fallbackTextFromPdfFormulaLatex,
  normalizePdfFormulaLatexSource,
} from "@mrnote/core/web/lib/pdf/formula-latex.ts";
import {
  isRenderablePdfFormulaLatex,
  pdfFormulaBoxDrawsLatex,
  pdfFormulaLatexValidationError,
} from "@mrnote/core/web/lib/pdf/formula-render.ts";

const DEFAULT_FORMULA_CONTEXT_LIMIT = 24;

export interface PdfFormulaDraftContextEntry {
  key: string;
  id: string;
  page: number;
  latex: string;
  normalizedLatex?: string;
  text?: string;
  error?: string;
}

export interface PdfFormulaDraftContextResult {
  edits: TextBoxEditLike[];
  boxes: TextBoxEditLike[];
  appliedDrafts: PdfFormulaDraftContextEntry[];
  invalidDrafts: PdfFormulaDraftContextEntry[];
}

function editPageNumber(edit: TextBoxEditLike): number {
  const page = Number(edit.page);
  return (Number.isFinite(page) ? Math.max(0, Math.trunc(page)) : 0) + 1;
}

function sortTextBoxesForReading(
  boxes: ReadonlyArray<TextBoxEditLike>,
): TextBoxEditLike[] {
  return [...boxes].sort((a, b) => {
    const ay = Number(a.y);
    const by = Number(b.y);
    if (Number.isFinite(ay) && Number.isFinite(by) && ay !== by) {
      return ay - by;
    }
    const ax = Number(a.x);
    const bx = Number(b.x);
    if (Number.isFinite(ax) && Number.isFinite(bx) && ax !== bx) {
      return ax - bx;
    }
    return String(a.id || "").localeCompare(String(b.id || ""));
  });
}

function formulaContextLine(
  box: TextBoxEditLike,
  options: { allowExplicitFormulaKind?: boolean } = {},
): string | null {
  const text = String(box.text ?? "").trim();
  const latex = String(box.latex ?? "").trim();
  if (!isFormulaTextBox(box, options)) return null;
  if (!text && !latex) return null;
  const id = String(box.id || "").trim();
  const label = id ? `[formula id=${id}]` : "[formula]";
  if (text && latex && text !== latex) return `${label} ${text}\nLaTeX: ${latex}`;
  if (text) return `${label} ${text}`;
  return `${label}\nLaTeX: ${latex}`;
}

function pdfFormulaLatexQualityError(latex: string): string | null {
  const validationError = pdfFormulaLatexValidationError(latex);
  if (validationError) return validationError;
  return isRenderablePdfFormulaLatex(latex)
    ? null
    : "Formula source is too weak to treat as a structured formula";
}

export function isPdfFormulaTextBoxContextCandidate(
  box: TextBoxEditLike,
): boolean {
  if (box.kind === "text") return false;
  if (!isFormulaTextBoxKind(box)) return false;
  const latex = String(box.latex ?? "").trim();
  if (!latex || !pdfFormulaBoxDrawsLatex(box, latex)) return false;
  if (box.kind === "formula") return true;
  const text = String(box.text ?? "").trim();
  if (!text) return true;
  return formulaTextLooksCompatibleWithLatex(text, latex);
}

function isFormulaTextBox(
  box: TextBoxEditLike,
  options: { allowExplicitFormulaKind?: boolean } = {},
): boolean {
  if (options.allowExplicitFormulaKind && box.kind === "formula") {
    const latex = String(box.latex ?? "").trim();
    return !latex || pdfFormulaLatexValidationError(latex) == null;
  }
  if (box.kind === "formula") return isPdfFormulaTextBoxContextCandidate(box);
  if (box.kind === "text") return false;
  return isPdfFormulaTextBoxContextCandidate(box);
}

function parseFormulaDraftKey(
  key: string,
): { page: number; id: string } | null {
  const split = key.indexOf(":");
  if (split <= 0) return null;
  const page = Number(key.slice(0, split));
  const id = key.slice(split + 1).trim();
  if (!Number.isInteger(page) || page < 0 || !id) return null;
  return { page, id };
}

function textBoxWithFormulaDraft<T extends TextBoxEditLike>(
  box: T,
  normalizedLatex: string,
): T & { draft?: boolean } {
  return {
    ...box,
    kind: "formula",
    latex: normalizedLatex,
    text: fallbackTextFromPdfFormulaLatex(normalizedLatex),
    draft: true,
  };
}

export function applyPdfFormulaDraftsToContext({
  edits,
  visibleTextBoxes,
  formulaDraftsByKey,
  pageNumber,
}: {
  edits: ReadonlyArray<TextBoxEditLike> | null | undefined;
  visibleTextBoxes: ReadonlyArray<TextBoxEditLike> | null | undefined;
  formulaDraftsByKey: Record<string, string> | null | undefined;
  pageNumber: number;
}): PdfFormulaDraftContextResult {
  const pageIndex = Math.max(0, Math.trunc(pageNumber) - 1);
  const nextEdits = Array.isArray(edits) ? [...edits] : [];
  const nextBoxes = Array.isArray(visibleTextBoxes) ? [...visibleTextBoxes] : [];
  const editsById = new Map(
    nextEdits
      .filter((edit) => editPageNumber(edit) === pageNumber)
      .map((edit, index) => [String(edit.id || ""), { edit, index }]),
  );
  const boxesById = new Map(
    nextBoxes
      .filter((box) => editPageNumber(box) === pageNumber)
      .map((box, index) => [String(box.id || ""), { box, index }]),
  );
  const appliedDrafts: PdfFormulaDraftContextEntry[] = [];
  const invalidDrafts: PdfFormulaDraftContextEntry[] = [];

  for (const [key, latex] of Object.entries(formulaDraftsByKey ?? {})) {
    const parsed = parseFormulaDraftKey(key);
    if (!parsed || parsed.page !== pageIndex) continue;
    const editEntry = editsById.get(parsed.id);
    const boxEntry = boxesById.get(parsed.id);
    const base = editEntry?.edit ?? boxEntry?.box;
    if (!base || !isFormulaTextBox(base, { allowExplicitFormulaKind: true })) {
      continue;
    }

    const normalizedLatex = normalizePdfFormulaLatexSource(latex);
    const error = pdfFormulaLatexQualityError(normalizedLatex);
    if (error) {
      invalidDrafts.push({
        key,
        id: parsed.id,
        page: parsed.page,
        latex,
        normalizedLatex,
        error,
      });
      continue;
    }

    const entry: PdfFormulaDraftContextEntry = {
      key,
      id: parsed.id,
      page: parsed.page,
      latex,
      normalizedLatex,
      text: fallbackTextFromPdfFormulaLatex(normalizedLatex),
    };
    appliedDrafts.push(entry);

    const editDraft = textBoxWithFormulaDraft(base, normalizedLatex);
    if (editEntry) {
      nextEdits[editEntry.index] = editDraft;
    } else {
      nextEdits.push(editDraft);
    }

    const boxDraft = textBoxWithFormulaDraft(
      boxEntry?.box ?? editDraft,
      normalizedLatex,
    );
    if (boxEntry) {
      nextBoxes[boxEntry.index] = boxDraft;
    } else {
      nextBoxes.push(boxDraft);
    }
  }

  return {
    edits: nextEdits,
    boxes: nextBoxes,
    appliedDrafts,
    invalidDrafts,
  };
}

function normalizeContextWhitespace(text: string): string {
  return text
    .split("\n")
    .map((line) => line.replace(/[ \t]{2,}/g, " ").trim())
    .filter(Boolean)
    .join("\n")
    .trim();
}

const FORMULA_GREEK_TO_ASCII: Record<string, string> = {
  α: "alpha",
  β: "beta",
  γ: "gamma",
  δ: "delta",
  Δ: "Delta",
  ε: "epsilon",
  θ: "theta",
  λ: "lambda",
  μ: "mu",
  π: "pi",
  σ: "sigma",
  Σ: "Sigma",
  φ: "phi",
  Φ: "Phi",
  ω: "omega",
  Ω: "Omega",
};

const FORMULA_SUPERSCRIPT_DIGITS: Record<string, string> = {
  "0": "⁰",
  "1": "¹",
  "2": "²",
  "3": "³",
  "4": "⁴",
  "5": "⁵",
  "6": "⁶",
  "7": "⁷",
  "8": "⁸",
  "9": "⁹",
};

const FORMULA_SUPERSCRIPT_TO_ASCII = Object.fromEntries(
  Object.entries(FORMULA_SUPERSCRIPT_DIGITS).map(([digit, superscript]) => [
    superscript,
    digit,
  ]),
);

function formulaTextVariants(text: string): string[] {
  const variants = new Set<string>();
  const add = (value: string) => {
    const normalized = normalizeContextWhitespace(value);
    if (normalized) variants.add(normalized);
  };
  add(text);
  add(
    text.replace(
      /[αβγδεθλμπσΣφΦωΩΔ]/g,
      (char) => FORMULA_GREEK_TO_ASCII[char] ?? char,
    ),
  );
  add(
    text.replace(
      /\b(alpha|beta|gamma|delta|Delta|epsilon|theta|lambda|mu|pi|sigma|Sigma|phi|Phi|omega|Omega)\b/g,
      (name) => {
        const entry = Object.entries(FORMULA_GREEK_TO_ASCII).find(
          ([, ascii]) => ascii === name,
        );
        return entry?.[0] ?? name;
      },
    ),
  );
  add(
    text.replace(
      /([A-Za-z0-9)\]])([⁰¹²³⁴⁵⁶⁷⁸⁹]+)/g,
      (_, base: string, superscript: string) =>
        `${base}^${Array.from(superscript)
          .map((char) => FORMULA_SUPERSCRIPT_TO_ASCII[char] ?? char)
          .join("")}`,
    ),
  );
  add(
    text.replace(/\^([0-9]+)/g, (_, digits: string) =>
      Array.from(digits)
        .map((digit) => FORMULA_SUPERSCRIPT_DIGITS[digit] ?? digit)
        .join(""),
    ),
  );
  return [...variants];
}

function compactFormulaContextText(text: string): string {
  return normalizeContextWhitespace(text).replace(/\s+/g, "");
}

function formulaTextLooksCompatibleWithLatex(
  text: string,
  latex: string,
): boolean {
  const textVariants = formulaTextVariants(text).map(compactFormulaContextText);
  const latexFallbackVariants = formulaTextVariants(
    fallbackTextFromPdfFormulaLatex(latex),
  ).map(compactFormulaContextText);
  return textVariants.some(
    (variant) => variant && latexFallbackVariants.includes(variant),
  );
}

function formulaRemovalTexts(box: TextBoxEditLike): string[] {
  const texts = [String(box.text ?? "").trim()];
  const latex = String(box.latex ?? "").trim();
  if (latex) texts.push(fallbackTextFromPdfFormulaLatex(latex));
  return Array.from(new Set(texts.flatMap(formulaTextVariants)));
}

function flexibleTextPattern(text: string): RegExp | null {
  const trimmed = text.trim();
  if (!trimmed || trimmed.length < 4) return null;
  const parts = trimmed
    .split(/\s+/)
    .map((part) => part.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"));
  if (parts.length === 0) return null;
  return new RegExp(parts.join("\\s+"), "g");
}

function shortFormulaTextPattern(text: string): RegExp | null {
  const trimmed = text.trim();
  if (!trimmed || trimmed.length >= 4) return null;
  const escaped = trimmed.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return new RegExp(`(^|[^\\p{L}\\p{N}_])${escaped}(?=$|[^\\p{L}\\p{N}_])`, "gu");
}

function removeFormulaTextFromRawContext(rawText: string, text: string): string {
  const flexiblePattern = flexibleTextPattern(text);
  if (flexiblePattern) return rawText.replace(flexiblePattern, " ");
  const shortPattern = shortFormulaTextPattern(text);
  if (!shortPattern) return rawText;
  return rawText.replace(shortPattern, (_match, prefix: string) =>
    prefix ? `${prefix} ` : " ",
  );
}

function rawTextWithOverriddenBoxesRemoved(
  pageText: string,
  edits: ReadonlyArray<TextBoxEditLike> | null | undefined,
  visibleTextBoxes: ReadonlyArray<TextBoxEditLike> | null | undefined,
  pageNumber: number,
  formulaRemovalIds: ReadonlySet<string> | null = null,
): string {
  const rawText = pageText.trim();
  if (!rawText || !Array.isArray(visibleTextBoxes)) {
    return rawText;
  }
  const visibleById = new Map(
    visibleTextBoxes
      .filter((box) => editPageNumber(box) === pageNumber)
      .map((box) => [String(box.id || ""), box]),
  );
  const overriddenTexts = new Set<string>();
  if (Array.isArray(edits)) {
    for (const edit of edits) {
      if (editPageNumber(edit) !== pageNumber) continue;
      const id = String(edit.id || "");
      const original = visibleById.get(id);
      if (!original) continue;
      const originalText = String(original.text ?? "").trim();
      const editedText = String(edit.text ?? "").trim();
      if (
        originalText &&
        (originalText !== editedText ||
          isFormulaTextBox(edit, { allowExplicitFormulaKind: true }) ||
          isFormulaTextBox(original))
      ) {
        for (const text of formulaRemovalTexts(original)) {
          overriddenTexts.add(text);
        }
      }
    }
  }
  for (const box of visibleById.values()) {
    const id = String(box.id || "");
    if (formulaRemovalIds && !formulaRemovalIds.has(id)) continue;
    if (!isFormulaTextBox(box)) continue;
    for (const text of formulaRemovalTexts(box)) {
      overriddenTexts.add(text);
    }
  }
  if (overriddenTexts.size === 0) return rawText;

  let nextText = rawText;
  for (const text of [...overriddenTexts].sort((a, b) => b.length - a.length)) {
    nextText = removeFormulaTextFromRawContext(nextText, text);
  }
  return normalizeContextWhitespace(nextText);
}

function selectedFormulaTextBoxesForContext(
  boxes: ReadonlyArray<TextBoxEditLike> | null | undefined,
  pageNumber: number,
  excludedIds: ReadonlySet<string> = new Set(),
): TextBoxEditLike[] {
  if (!Array.isArray(boxes) || boxes.length === 0) return [];
  return sortTextBoxesForReading(
    boxes.filter((box) => {
      const id = String(box.id || "");
      return editPageNumber(box) === pageNumber && !excludedIds.has(id);
    }),
  )
    .filter((box) => isFormulaTextBox(box))
    .slice(0, DEFAULT_FORMULA_CONTEXT_LIMIT);
}

export function textBoxEditContextForPage(
  edits: ReadonlyArray<TextBoxEditLike> | null | undefined,
  pageNumber: number,
): string {
  if (!Array.isArray(edits) || edits.length === 0) return "";
  const lines = sortTextBoxesForReading(
    edits.filter((edit) => editPageNumber(edit) === pageNumber),
  )
    .map((edit) => {
      const formulaLine = formulaContextLine(edit);
      if (formulaLine) return formulaLine;
      const text = String(edit.text ?? "").trim();
      if (text) return text;
      if (edit.created) return null;
      const id = String(edit.id || "").trim();
      return id ? `[cleared text box: ${id}]` : "[cleared text box]";
    })
    .filter((line): line is string => Boolean(line));
  if (lines.length === 0) return "";
  return [
    `PDF edited text on page ${pageNumber} (overrides matching extracted text boxes):`,
    ...lines,
  ].join("\n");
}

export function textBoxFormulaContextForPage(
  boxes: ReadonlyArray<TextBoxEditLike> | null | undefined,
  pageNumber: number,
  excludedIds: ReadonlySet<string> = new Set(),
): string {
  const lines = selectedFormulaTextBoxesForContext(
    boxes,
    pageNumber,
    excludedIds,
  )
    .map((box) => formulaContextLine(box))
    .filter((line): line is string => Boolean(line));
  if (lines.length === 0) return "";
  return [
    `PDF formula text boxes on page ${pageNumber} (visible formulas with LaTeX where available):`,
    ...lines,
  ].join("\n");
}

export function mergePdfPageTextWithTextBoxEdits(
  pageText: string,
  edits: ReadonlyArray<TextBoxEditLike> | null | undefined,
  pageNumber: number,
  visibleTextBoxes?: ReadonlyArray<TextBoxEditLike> | null,
): string {
  const editedText = textBoxEditContextForPage(edits, pageNumber);
  const editedIds = new Set(
    Array.isArray(edits)
      ? edits
          .filter((edit) => editPageNumber(edit) === pageNumber)
          .map((edit) => String(edit.id || ""))
      : [],
  );
  const formulaText = textBoxFormulaContextForPage(
    visibleTextBoxes,
    pageNumber,
    editedIds,
  );
  const selectedFormulaIds = new Set(
    selectedFormulaTextBoxesForContext(
      visibleTextBoxes,
      pageNumber,
      editedIds,
    ).map((box) => String(box.id || "")),
  );
  const rawText = rawTextWithOverriddenBoxesRemoved(
    pageText,
    edits,
    visibleTextBoxes,
    pageNumber,
    selectedFormulaIds,
  );
  const sections = [
    editedText,
    formulaText,
    rawText ? `PDF extracted text on page ${pageNumber}:\n${rawText}` : "",
  ].filter(Boolean);
  if (sections.length === 0) return "";
  if (!editedText && !formulaText) return rawText;
  return sections.join("\n\n");
}
