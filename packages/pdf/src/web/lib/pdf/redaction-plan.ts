/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Advanced redaction (WPS PDF ▸ 密文/标记密文).
 *
 * Redaction has one failure mode that matters: drawing a black rectangle over
 * text and shipping the file with the text still in it. Every real redaction
 * tool therefore separates MARKING from APPLYING, and verifies afterwards that
 * the words are gone from the content — not merely covered.
 *
 * This module is the marking and the verification. It finds what a user asked
 * to redact (a pattern, a search term, whole pages), and after the host has
 * applied the removal it re-reads the extracted text and says plainly whether
 * the redaction actually took. A plan that cannot be verified is reported as
 * unverified rather than described as done.
 */

export type RedactionPatternKind =
  | "email"
  | "phone"
  | "credit-card"
  | "id-number"
  | "custom";

export interface RedactionMark {
  /** 1-based page number. */
  page: number;
  /** The exact text this mark covers. */
  text: string;
  /** Offset of the text within that page's extracted string. */
  start: number;
  kind: RedactionPatternKind;
}

export interface RedactionRequest {
  /** Extracted text per page, index 0 = page 1. */
  pages: readonly string[];
  patterns?: readonly RedactionPatternKind[];
  /** Literal terms to redact, matched case-insensitively. */
  terms?: readonly string[];
  /** 1-based pages to redact in full. */
  wholePages?: readonly number[];
  /** A user-supplied regular expression source, used with `custom`. */
  customPattern?: string;
}

export interface RedactionPlan {
  marks: RedactionMark[];
  /** Pages to be emptied entirely. */
  wholePages: number[];
  /** A custom pattern that would not compile, so nothing was searched for it. */
  invalidCustomPattern: boolean;
}

const PATTERNS: Record<Exclude<RedactionPatternKind, "custom">, RegExp> = {
  email: /[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g,
  // Long enough to be a phone number rather than a year or a quantity.
  phone: /(?:\+?\d[\d\s-]{7,}\d)/g,
  "credit-card": /\b(?:\d[ -]?){13,19}\b/g,
  // Mainland ID: 17 digits plus a check character.
  "id-number": /\b\d{17}[\dXx]\b/g,
};

export const MAX_REDACTION_MARKS = 10_000;

/**
 * A number-shaped match is only a phone number or a card number when its DIGIT
 * COUNT fits, and when it is not simply a range.
 *
 * Over-matching is the worse failure for redaction: blacking out the year
 * ranges and reference numbers in a report destroys the document, and a tool
 * that does that gets switched off — after which nothing is redacted at all.
 */
function isPlausibleNumber(
  candidate: string,
  kind: "phone" | "credit-card",
): boolean {
  const digits = (candidate.match(/\d/g) ?? []).length;
  // A range is two runs of the SAME length, joined by one hyphen, ascending:
  // `1999-2024`, `1000000-2000000`. Requiring all three keeps a real phone
  // number written with a single hyphen — `0207-9460958`, whose halves differ
  // in length — from being dismissed as a range and left un-redacted, which is
  // the failure the over-matching fix could otherwise have introduced.
  const range = /^(\d+)-(\d+)$/.exec(candidate.trim());
  if (
    range &&
    range[1].length === range[2].length &&
    Number(range[1]) < Number(range[2])
  ) {
    return false;
  }
  return kind === "phone"
    ? digits >= 9 && digits <= 15 // E.164 allows at most 15.
    : digits >= 13 && digits <= 19;
}

/** Find everything the request asks to redact. */
export function planRedaction(request: RedactionRequest): RedactionPlan {
  const marks: RedactionMark[] = [];
  let invalidCustomPattern = false;

  let custom: RegExp | null = null;
  if (request.patterns?.includes("custom") && request.customPattern) {
    try {
      custom = new RegExp(request.customPattern, "g");
    } catch {
      invalidCustomPattern = true;
    }
  }

  request.pages.forEach((text, index) => {
    const page = index + 1;
    const push = (value: string, start: number, kind: RedactionPatternKind) => {
      if (marks.length >= MAX_REDACTION_MARKS || !value) return;
      marks.push({ page, text: value, start, kind });
    };

    for (const kind of request.patterns ?? []) {
      if (kind === "custom") {
        if (!custom) continue;
        custom.lastIndex = 0;
        for (const match of text.matchAll(custom)) {
          push(match[0], match.index ?? 0, "custom");
        }
        continue;
      }
      const pattern = PATTERNS[kind];
      pattern.lastIndex = 0;
      for (const match of text.matchAll(pattern)) {
        if (
          (kind === "phone" || kind === "credit-card") &&
          !isPlausibleNumber(match[0], kind)
        ) {
          continue;
        }
        push(match[0], match.index ?? 0, kind);
      }
    }

    for (const term of request.terms ?? []) {
      if (!term) continue;
      // Matched against the ORIGINAL text, not a lower-cased copy of it. Case
      // folding is not length-preserving — U+0130 lower-cases to two code
      // points — so an index found in the copy did not point at the same place
      // in the text it was then sliced from. One Turkish dotted capital earlier
      // on the page moved the mark for "SECRET" one character left: it covered
      // "ECRET " and left the S sitting on the page. Text that stays visible is
      // the one failure a redaction tool cannot have.
      //
      // The cost is that `i` uses simple case folding, so a search for
      // "istanbul" no longer reaches "İstanbul". Missing that is better than
      // covering the wrong characters, which fails silently and looks redacted.
      let matcher: RegExp;
      try {
        matcher = new RegExp(term.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "giu");
      } catch {
        continue;
      }
      for (const match of text.matchAll(matcher)) {
        push(match[0], match.index ?? 0, "custom");
      }
    }
  });

  const wholePages = [...new Set(request.wholePages ?? [])]
    .filter((page) => page >= 1 && page <= request.pages.length)
    .sort((a, b) => a - b);

  return {
    marks: marks.sort((a, b) => a.page - b.page || a.start - b.start),
    wholePages,
    invalidCustomPattern,
  };
}

export interface RedactionVerification {
  /** True only when every marked string is gone from the applied text. */
  verified: boolean;
  /** Marks whose text is still present — the black-rectangle failure. */
  survivingMarks: RedactionMark[];
  /** Pages that were meant to be emptied but still carry text. */
  survivingPages: number[];
}

/**
 * Re-read the applied document and say whether the redaction took. This is the
 * check that separates a redaction from a black rectangle, so it is deliberately
 * strict: any surviving occurrence fails the whole verification.
 */
export function verifyRedaction(
  plan: RedactionPlan,
  appliedPages: readonly string[],
): RedactionVerification {
  const survivingMarks = plan.marks.filter((mark) => {
    const text = appliedPages[mark.page - 1] ?? "";
    return text.toLowerCase().includes(mark.text.toLowerCase());
  });
  const survivingPages = plan.wholePages.filter(
    (page) => (appliedPages[page - 1] ?? "").trim() !== "",
  );
  return {
    verified: survivingMarks.length === 0 && survivingPages.length === 0,
    survivingMarks,
    survivingPages,
  };
}
