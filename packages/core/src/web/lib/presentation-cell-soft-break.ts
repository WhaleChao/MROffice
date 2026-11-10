/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A table cell's line break inside a paragraph: PowerPoint's Shift+Enter,
 * DrawingML `a:br` (ECMA-376 Part 1 §21.1.2.2.1). The page keeps it as U+000B
 * in the cell's text, as python-pptx reads a cell; Enter (`\n`) starts a new
 * paragraph (`a:p`). CSS breaks a line only at `\n`, so a break is drawn as
 * one.
 */
export const CELL_SOFT_BREAK = "\v";

/** The cell text as drawn: each line break a CSS line break. */
export function cellDisplayText(text: string): string {
  return text.includes(CELL_SOFT_BREAK) ? text.replace(/\v/g, "\n") : text;
}
