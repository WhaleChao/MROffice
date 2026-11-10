/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
const SCRIPT_SIZE = 2 / 3;
const SUPERSCRIPT_RISE = 1 / 3;
const SUBSCRIPT_RISE = -1 / 6;

export type PdfNativeTextRiseControl = "superscript" | "subscript";

export interface PdfNativeTextRiseState {
  readonly fontSize: number;
  readonly textRise: number;
}

function round(value: number): number {
  return Math.round(value * 100) / 100;
}

/** The size the run was scripted from, when it carries this convention's rise. */
function conventionalScriptBase(original: PdfNativeTextRiseState): number | null {
  const { fontSize, textRise } = original;
  if (!(fontSize > 0) || textRise === 0) return null;
  const base = fontSize / SCRIPT_SIZE;
  const riseRatio = textRise > 0 ? SUPERSCRIPT_RISE : SUBSCRIPT_RISE;
  const expectedRise = base * riseRatio;
  // Both values were written rounded to 0.01.
  if (Math.abs(textRise - expectedRise) > 0.02 + base * 0.01) return null;
  // Dividing the rounded size back gives 22.005 for a 22 pt run (14.67 / ⅔);
  // the coarsest size that writes the same rounded script size and rise is
  // the one it was scripted from, so turning the script off restores 22, not
  // 22.01.
  for (const step of [1, 0.5, 0.25, 0.1, 0.01]) {
    const candidate = Math.round(base / step) * step;
    if (
      round(candidate * SCRIPT_SIZE) === fontSize &&
      Math.abs(round(candidate * riseRatio) - textRise) <= 0.011
    ) {
      return candidate;
    }
  }
  return base;
}

/** The run at its normal baseline, as far as the source lets us know it. */
export function nativeTextNormalState(
  original: PdfNativeTextRiseState,
): PdfNativeTextRiseState {
  if (original.textRise === 0) return original;
  const base = conventionalScriptBase(original);
  return { fontSize: base === null ? original.fontSize : round(base), textRise: 0 };
}

export function nativeTextRiseToggle(
  original: PdfNativeTextRiseState,
  currentRise: number,
  control: PdfNativeTextRiseControl,
): PdfNativeTextRiseState {
  const active = control === "superscript" ? currentRise > 0 : currentRise < 0;
  const originalActive =
    control === "superscript" ? original.textRise > 0 : original.textRise < 0;
  if (active) {
    // Added in this edit: put the run back exactly as it was opened.
    if (!originalActive) return original;
    return nativeTextNormalState(original);
  }
  const base = nativeTextNormalState(original).fontSize;
  return {
    fontSize: round(base * SCRIPT_SIZE),
    textRise: round(base * (control === "superscript" ? SUPERSCRIPT_RISE : SUBSCRIPT_RISE)),
  };
}







export function nativeTextSizeOfScriptedRun(asked: number, currentRise: number): PdfNativeTextRiseState {
  if (currentRise === 0) return { fontSize: asked, textRise: 0 };
  return {
    fontSize: round(asked * SCRIPT_SIZE),
    textRise: round(asked * (currentRise > 0 ? SUPERSCRIPT_RISE : SUBSCRIPT_RISE)),
  };
}

/** A script's size, of the size it is scripted from (Word). */
export const PDF_NATIVE_SCRIPT_SIZE = SCRIPT_SIZE;

/** A script's rise in ems of its own, scripted size: a superscript raised a
 * third of the size it is scripted from (half its own), a subscript lowered
 * a sixth (a quarter). The box formats the window plans and the draft's CSS
 * read it here, as the presses above write it. */
export function pdfNativeScriptRiseEm(script: PdfNativeTextRiseControl): number {
  return (script === "superscript" ? SUPERSCRIPT_RISE : SUBSCRIPT_RISE) / SCRIPT_SIZE;
}

/** The factor a run's size takes when `script` is set on a run standing as
 * `standing`: 2/3 onto its baseline's size, back to the size it was scripted
 * from when the script comes off -- the size nativeTextNormalState reads of
 * `own` (the run as its proof read it) where given, so 14.67 comes back to
 * 22, not to the 22.005 of a bare 3/2 -- and 1 from one script to the other. */
export function pdfNativeScriptSizeFactor(
  standing: PdfNativeTextRiseControl | "none",
  script: PdfNativeTextRiseControl | "none",
  own?: PdfNativeTextRiseState,
): number {
  if (script === standing) return 1;
  if (standing === "none") return SCRIPT_SIZE;
  if (script !== "none") return 1;
  if (own && own.fontSize > 0 && own.textRise !== 0 && (own.textRise > 0) === (standing === "superscript")) {
    return nativeTextNormalState(own).fontSize / own.fontSize;
  }
  return 1 / SCRIPT_SIZE;
}
