/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * An equation box's size, text insets and usable-text check, apart from the
 * slide commands in presentation-equation-buttons (which re-exports them):
 * the spreadsheet's equation box needs only these, and the commands module
 * brought the deck's shape, slide and copy commands with it.
 */


/** An equation as the server answered it for the page. */
export interface PresentationEquation {
  /** The characters the equation draws in the line (the run's `text`). */
  readonly text: string;
  readonly latex: string;
  /** Stands as its own line (`m:oMathPara`). */
  readonly display?: boolean;
}

/** The drawn characters an equation run can hold (an `a14:m` has no line break). */
export function presentationEquationTextIsUsable(text: string): boolean {
  return text.length > 0 && !/[\n\r\v]/.test(text);
}

/**
 * PowerPoint's text box insets (`a:bodyPr` defaults: 0.1" left and right,
 * 0.05" top and bottom), in EMU. An equation box states them, so the editor
 * draws exactly the frame the file has.
 */
export const EQUATION_BOX_TEXT_INSETS = { left: 91_440, top: 45_720, right: 91_440, bottom: 45_720 } as const;

/** EMU per CSS px on the slide's design surface (914400 EMU/in at 96 px/in). */
export const EMU_PER_PX = 9_525;

export type Insets = { readonly left: number; readonly top: number; readonly right: number; readonly bottom: number };

/**
 * The size (EMU) an equation box needs for its equation as drawn
 * (`contentPx`, design px) plus its insets: "Resize shape to fit text" with
 * no wrap, as PowerPoint's Insert > Equation box is.
 */
export function equationBoxSizeFor(
  contentPx: { readonly width: number; readonly height: number },
  insets: Insets = EQUATION_BOX_TEXT_INSETS,
): { readonly width: number; readonly height: number } {
  return {
    width: Math.ceil(contentPx.width) * EMU_PER_PX + insets.left + insets.right,
    height: Math.ceil(contentPx.height) * EMU_PER_PX + insets.top + insets.bottom,
  };
}
