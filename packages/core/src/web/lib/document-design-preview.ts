/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * How a Quick Style looks under a design, as gallery thumbnails draw it:
 * Word's Styles gallery and its style-set gallery both draw each style in
 * the document's own definition (style set plus theme), so their previews
 * follow every Design change.
 */

import type { CSSProperties } from "react";

import {
  documentDesignBase,
  documentDesignStyleRecipe,
  type DocumentDesign,
  type DocumentStyleRecipe,
} from "./document-design";
import type { WordQuickStyleId } from "./office-word-style-sets";

function quoted(face: string): string {
  return `"${face.replace(/["'\\;{}()]/g, "").slice(0, 200)}"`;
}

function faceStack(...faces: (string | null | undefined)[]): string {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const face of faces) {
    const trimmed = face?.trim();
    if (!trimmed || seen.has(trimmed.toLowerCase())) continue;
    seen.add(trimmed.toLowerCase());
    out.push(quoted(trimmed));
  }
  return [...out, '"PingFang SC"', '"Microsoft YaHei"', "sans-serif"].join(", ");
}

function px(value: string | null): number | null {
  const parsed = value ? Number.parseFloat(value) : NaN;
  return Number.isFinite(parsed) ? parsed : null;
}

/**
 * The CSS a preview of one style draws with, its size scaled by `scale`
 * (1 = the document's own size) and kept between `minPx` and `maxPx`.
 * Undefined while Word's style sets are not here (the caller's fallback draws).
 */
export function documentDesignStylePreviewCss(
  design: DocumentDesign,
  id: WordQuickStyleId,
  { scale = 1, minPx = 7, maxPx = 40 }: { scale?: number; minPx?: number; maxPx?: number } = {},
): CSSProperties | undefined {
  const recipe: DocumentStyleRecipe | null = documentDesignStyleRecipe(design, id);
  const base = documentDesignBase(design);
  if (!recipe || !base) return undefined;
  const { run, paragraph } = recipe;
  const size = (px(run.fontSize) ?? px(base.defaultFontSize) ?? 14.667) * scale;
  const css: CSSProperties = {
    fontFamily: run.fontFamily || run.fontCjkFamily
      ? faceStack(run.fontFamily, run.fontCjkFamily)
      : base.defaultFontFamily,
    fontSize: `${Math.round(Math.min(maxPx, Math.max(minPx, size)) * 100) / 100}px`,
    color: run.textColor ?? base.defaultTextColor ?? "#000000",
    fontWeight: run.bold ? 700 : 400,
    fontStyle: run.italic ? "italic" : "normal",
    fontVariant: run.smallCaps ? "small-caps" : "normal",
    textTransform: run.allCaps ? "uppercase" : "none",
    letterSpacing: run.charSpacingPt ? `${run.charSpacingPt * scale}pt` : "normal",
    textDecoration: run.underline ? "underline" : "none",
    textAlign: paragraph.textAlign ?? base.defaultTextAlign ?? "left",
  };
  if (paragraph.shading) css.backgroundColor = paragraph.shading;
  const top = paragraph.borders?.top;
  const bottom = paragraph.borders?.bottom;
  const left = paragraph.borders?.left;
  if (top?.color) css.borderTop = `1px solid ${top.color}`;
  if (bottom?.color) css.borderBottom = `1px solid ${bottom.color}`;
  if (left?.color) css.borderLeft = `2px solid ${left.color}`;
  return css;
}
