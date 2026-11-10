/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Home ▸ Fill ▸ Across Worksheets… and Justify (XLS-A011), as pure plans.
 *
 * Microsoft Support, "Fill data automatically in worksheet cells" and the
 * Fill Across Worksheets dialog (All / Contents / Formats): the selected
 * range is copied to the same range of the other worksheets. Justify
 * ("Distribute the contents of a column across rows"): the text of the
 * range's left column is joined and re-wrapped into lines no wider than the
 * range, one line per row from the top; numbers and formulas refuse.
 */

import { eastAsianTextWidth } from "@mrnote/core/web/lib/spreadsheet/grid-sizing.ts";

export type FillAcrossMode = "all" | "contents" | "formats";

interface FillCell {
  value?: string;
  formula?: string;
  kind?: string;
  style?: unknown;
  [key: string]: unknown;
}

const CONTENT_KEYS = ["value", "formula", "kind", "runs"] as const;

/**
 * The target sheet's cells with the source range copied over them: All
 * (content and format), Contents (format of the target kept) or Formats
 * (content of the target kept).
 */
export function fillAcrossCells<C extends FillCell>(
  source: Readonly<Record<string, C>>,
  target: Readonly<Record<string, C>>,
  range: { r1: number; c1: number; r2: number; c2: number },
  mode: FillAcrossMode,
): Record<string, C> {
  const out: Record<string, C> = { ...target };
  for (let r = range.r1; r <= range.r2; r += 1) {
    for (let c = range.c1; c <= range.c2; c += 1) {
      const key = `${r}:${c}`;
      const from = source[key];
      const to = target[key];
      if (mode === "all") {
        if (from) out[key] = { ...from };
        else delete out[key];
        continue;
      }
      const next: Record<string, unknown> = { ...(to ?? {}) };
      if (mode === "contents") {
        for (const field of CONTENT_KEYS) {
          if (from && from[field] !== undefined) next[field] = from[field];
          else delete next[field];
        }
      } else if (from?.style !== undefined) {
        next.style = from.style;
      } else {
        delete next.style;
      }
      const empty =
        (next.value === undefined || next.value === "") &&
        next.formula === undefined &&
        next.style === undefined;
      if (empty) delete out[key];
      else out[key] = { value: "", ...next } as C;
    }
  }
  return out;
}

export type JustifyPlan =
  | { kind: "ok"; lines: string[] }
  | { kind: "numbers-or-formulas" }
  | { kind: "empty" };

/**
 * Justify's lines for the left-column texts of the selection: words are
 * kept whole where they fit, a word wider than the line (or East-Asian text,
 * which has no spaces) breaks at the character that would overflow.
 */
export function justifyLines(
  cells: ReadonlyArray<{ value?: string; formula?: string; kind?: string } | undefined>,
  widthPx: number,
  measure: (text: string) => number = (text) => eastAsianTextWidth(text) * 7,
): JustifyPlan {
  const texts: string[] = [];
  for (const cell of cells) {
    if (!cell) continue;
    const value = cell.value ?? "";
    if (cell.formula || value.startsWith("=") || cell.kind === "number" || (value.trim() !== "" && Number.isFinite(Number(value.trim())))) {
      return { kind: "numbers-or-formulas" };
    }
    if (value.trim()) texts.push(value.trim());
  }
  if (texts.length === 0) return { kind: "empty" };
  const tokens = texts.join(" ").split(/\s+/).filter(Boolean);
  const lines: string[] = [];
  let line = "";
  const push = () => {
    if (line) lines.push(line);
    line = "";
  };
  for (const token of tokens) {
    const candidate = line ? `${line} ${token}` : token;
    if (measure(candidate) <= widthPx) {
      line = candidate;
      continue;
    }
    if (line && measure(token) <= widthPx) {
      push();
      line = token;
      continue;
    }
    // A word wider than the line, or unspaced East-Asian text: break it.
    let rest = line ? ` ${token}` : token;
    while (rest) {
      let take = 0;
      for (let end = 1; end <= rest.length; end += 1) {
        if (measure(line + rest.slice(0, end)) > widthPx) break;
        take = end;
      }
      if (take === 0) {
        if (line) {
          push();
          rest = rest.trimStart();
          continue;
        }
        take = 1;
      }
      line += rest.slice(0, take);
      rest = rest.slice(take);
      if (rest) {
        push();
        rest = rest.trimStart();
      }
    }
  }
  push();
  return { kind: "ok", lines };
}
