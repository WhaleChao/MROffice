/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export type PdfTextBoxKind = "text" | "formula";

export interface TextBoxKindLike {
  kind?: PdfTextBoxKind;
  latex?: string;
  text?: string;
  /** A box the user made (插入 › 公式 makes one with no source yet). */
  created?: boolean;
}

export function hasFormulaLatex(latex: string | undefined): boolean {
  return typeof latex === "string" && latex.trim().length > 0;
}

export function inferTextBoxKind(
  source?: TextBoxKindLike,
  edit?: TextBoxKindLike,
): PdfTextBoxKind | undefined {
  if (edit?.kind) return edit.kind;
  if (hasFormulaLatex(edit?.latex)) return "formula";
  if (source?.kind) return source.kind;
  if (hasFormulaLatex(source?.latex)) return "formula";
  return undefined;
}

function hasTextBoxFormulaContent(value: TextBoxKindLike | undefined): boolean {
  if (!value) return false;
  if (hasFormulaLatex(value.latex)) return true;
  return typeof value.text === "string" && value.text.trim().length > 0;
}

export function hasUsableFormulaTextBoxContent(
  source?: TextBoxKindLike,
  edit?: TextBoxKindLike,
): boolean {
  return (
    inferTextBoxKind(source, edit) === "formula" &&
    (hasTextBoxFormulaContent(edit) || hasTextBoxFormulaContent(source))
  );
}

export function effectiveTextBoxKind(
  source?: TextBoxKindLike,
  edit?: TextBoxKindLike,
): PdfTextBoxKind | undefined {
  const kind = inferTextBoxKind(source, edit);
  if (kind !== "formula") return kind;
  // A formula the user has just made is one while its source is typed: an
  // empty one is dropped when it is left (as an empty new text box is).
  if (edit?.created === true || source?.created === true) return "formula";
  return hasUsableFormulaTextBoxContent(source, edit) ? "formula" : undefined;
}

export function latexForTextBoxKind(
  kind: PdfTextBoxKind | undefined,
  latex: string | undefined,
): string | undefined {
  return kind === "formula" ? latex : undefined;
}

export function isFormulaTextBoxKind(
  source?: TextBoxKindLike,
  edit?: TextBoxKindLike,
): boolean {
  return effectiveTextBoxKind(source, edit) === "formula";
}
