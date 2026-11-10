/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A box's rules (underline, strikethrough) asked as parts of its paragraph's
 * format: the paragraph rewrite draws them itself (paragraph-rewrite's rule
 * blocks) where the rows the paragraph is set in cannot each take their own
 * (a row of several runs). The format's parts do not overlap, so a part a
 * rule covers in part is cut where the rule starts and ends.
 */
import type {
  PdfNativeParagraphFormat,
  PdfNativeParagraphFormatFields,
} from "./engine/native/list-text-fidelity";

export interface PdfNativeParagraphRuleStretch {
  readonly style: "underline" | "strikethrough";
  readonly start: number;
  readonly end: number;
  readonly on: boolean;
}

type Part = PdfNativeParagraphFormatFields & Readonly<{ start: number; end: number }>;

export function pdfNativeParagraphFormatWithRules(
  format: PdfNativeParagraphFormat,
  rules: readonly PdfNativeParagraphRuleStretch[],
): PdfNativeParagraphFormat {
  const parts: readonly Part[] = format.parts ?? [];
  const cuts = new Set<number>();
  for (const part of parts) { cuts.add(part.start); cuts.add(part.end); }
  for (const rule of rules) if (rule.end > rule.start) { cuts.add(rule.start); cuts.add(rule.end); }
  const points = [...cuts].sort((left, right) => left - right);
  const next: Part[] = [];
  for (let index = 0; index + 1 < points.length; index += 1) {
    const start = points[index], end = points[index + 1];
    const own = parts.find((part) => part.start <= start && part.end >= end);
    const fields: { -readonly [K in keyof PdfNativeParagraphFormatFields]: PdfNativeParagraphFormatFields[K] } = {};
    if (own) {
      const { start: _start, end: _end, ...rest } = own;
      Object.assign(fields, rest);
    }
    for (const rule of rules) {
      if (rule.start <= start && rule.end >= end) fields[rule.style] = rule.on;
    }
    if (!Object.keys(fields).length) continue;
    // A piece set like the one before it is that piece, longer.
    const last = next.at(-1);
    if (last && last.end === start && JSON.stringify({ ...last, start: 0, end: 0 }) === JSON.stringify({ ...fields, start: 0, end: 0 })) {
      next[next.length - 1] = Object.freeze({ ...last, end });
      continue;
    }
    next.push(Object.freeze({ ...fields, start, end }));
  }
  const { parts: _parts, ...whole } = format;
  return Object.freeze({ ...whole, ...(next.length ? { parts: Object.freeze(next) } : {}) });
}
