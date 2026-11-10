/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export type PdfFormulaSourceTemplate =
  | "fraction"
  | "sqrt"
  | "superscript"
  | "subscript"
  | "sum"
  | "integral"
  | "parentheses"
  | "text";

export interface PdfFormulaSourceTemplateInput {
  source: string;
  selectionStart: number;
  selectionEnd: number;
  template: PdfFormulaSourceTemplate;
}

export interface PdfFormulaSourceTemplateResult {
  source: string;
  selectionStart: number;
  selectionEnd: number;
}

function clampSelectionOffset(value: number, length: number): number {
  if (!Number.isFinite(value)) return length;
  return Math.max(0, Math.min(length, Math.trunc(value)));
}

function insertionWithSelection({
  source,
  start,
  end,
  insertion,
  selectionStartInInsertion,
  selectionEndInInsertion = selectionStartInInsertion,
}: {
  source: string;
  start: number;
  end: number;
  insertion: string;
  selectionStartInInsertion: number;
  selectionEndInInsertion?: number;
}): PdfFormulaSourceTemplateResult {
  const nextSource = `${source.slice(0, start)}${insertion}${source.slice(end)}`;
  return {
    source: nextSource,
    selectionStart: start + selectionStartInInsertion,
    selectionEnd: start + selectionEndInInsertion,
  };
}

export function applyPdfFormulaSourceTemplate({
  source,
  selectionStart,
  selectionEnd,
  template,
}: PdfFormulaSourceTemplateInput): PdfFormulaSourceTemplateResult {
  const text = typeof source === "string" ? source : "";
  const from = clampSelectionOffset(selectionStart, text.length);
  const to = clampSelectionOffset(selectionEnd, text.length);
  const start = Math.min(from, to);
  const end = Math.max(from, to);
  const selected = text.slice(start, end);
  const selectedLength = selected.length;

  switch (template) {
    case "fraction": {
      const insertion = `\\frac{${selected}}{}`;
      const denominatorOffset = `\\frac{${selected}}{`.length;
      return insertionWithSelection({
        source: text,
        start,
        end,
        insertion,
        selectionStartInInsertion: selectedLength > 0 ? denominatorOffset : 6,
      });
    }
    case "sqrt": {
      const insertion = `\\sqrt{${selected}}`;
      return insertionWithSelection({
        source: text,
        start,
        end,
        insertion,
        selectionStartInInsertion: 6,
        selectionEndInInsertion: 6 + selectedLength,
      });
    }
    case "superscript": {
      const insertion = selected ? `${selected}^{}` : "^{}";
      return insertionWithSelection({
        source: text,
        start,
        end,
        insertion,
        selectionStartInInsertion: insertion.length - 1,
      });
    }
    case "subscript": {
      const insertion = selected ? `${selected}_{}` : "_{}";
      return insertionWithSelection({
        source: text,
        start,
        end,
        insertion,
        selectionStartInInsertion: insertion.length - 1,
      });
    }
    case "sum": {
      const insertion = selected ? `\\sum_{}^{} ${selected}` : "\\sum_{}^{}";
      return insertionWithSelection({
        source: text,
        start,
        end,
        insertion,
        selectionStartInInsertion: 6,
      });
    }
    case "integral": {
      const insertion = selected ? `\\int_{}^{} ${selected}` : "\\int_{}^{}";
      return insertionWithSelection({
        source: text,
        start,
        end,
        insertion,
        selectionStartInInsertion: 6,
      });
    }
    case "parentheses": {
      const insertion = `\\left(${selected}\\right)`;
      return insertionWithSelection({
        source: text,
        start,
        end,
        insertion,
        selectionStartInInsertion: 6,
        selectionEndInInsertion: 6 + selectedLength,
      });
    }
    case "text": {
      const insertion = `\\text{${selected}}`;
      return insertionWithSelection({
        source: text,
        start,
        end,
        insertion,
        selectionStartInInsertion: 6,
        selectionEndInInsertion: 6 + selectedLength,
      });
    }
  }
}
