/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Smart reference jumping index for the PDF reader (STEM cross references).
 *
 * A STEM PDF is full of cross references — "Eq. (1.2)", "式(3.1)", "Fig. 4",
 * "图 2.8", "Theorem 3.1", "[12]" — whose TARGETS (the numbered equation line,
 * the figure/table caption, the theorem-like heading, the numbered section
 * heading, the bibliography entry) live elsewhere in the document. This module
 * builds a pure index over the same per-page text items the find pipeline
 * consumes (`buildSearchablePageText` from find-normalize.ts):
 *
 *   TARGETS  — pattern-scanned per raw-text LINE (line structure comes from
 *              the find-normalize page assembly: `hasEOL` / vertical jumps →
 *              "\n"). The equation-number-at-line-end pattern is adapted from
 *              `isEquationNumberLine` in text-boxes.ts (the formula pipeline's
 *              recognizer), narrowed to the paren form. The FIRST occurrence
 *              of a key in document order wins.
 *   MENTIONS — keyword+number phrases scanned across each page's raw text,
 *              kept ONLY when a matching target exists (precision over recall)
 *              and the phrase is not a target's own marker (a caption must not
 *              link to itself).
 *
 * Mention ranges are returned in FOLDED-text offsets — the same coordinate
 * space as find matches — so the viewer reuses the find pipeline's span-rect
 * measurement unchanged. Targets carry the PDF user-space top coordinate of
 * their line (from the source item transforms) so navigation can reuse the
 * /XYZ dest-offset path (`destTopToPageOffset`), plus a `yFraction` for
 * consumers without user-space geometry.
 *
 * Everything here is pure (no DOM, no pdf.js) so it unit-tests without jsdom.
 */

import {
  buildSearchablePageText,
  type PdfPageSearchText,
  type PdfSearchTextItem,
} from "./find-normalize";
import { itemsForRange } from "@mrnote/core/web/lib/pdf/text-search.ts";

/** One page's input: the find-pipeline text items plus page geometry. */
export interface PdfReferencePageInput {
  /** pdf.js text items (or plain strings — those carry no geometry). */
  items: ReadonlyArray<PdfSearchTextItem | string>;
  /** PDF user-space page height (points at scale 1), for target offsets. */
  pageHeight?: number;
}

/** A jump target (caption / equation number / heading / bib entry). */
export interface PdfReferenceTarget {
  /** Canonical key, e.g. "eq:1.2", "fig:2.8", "thm:3.1", "bib:12". */
  key: string;
  /** Index into the input pages array (the viewer's visible-page order). */
  pageIndex: number;
  /**
   * PDF user-space top coordinate of the target line (like a /FitH top), or
   * null when the source items carry no geometry (OCR / plain-string pages).
   */
  top: number | null;
  /** 0 = page top … 1 = page bottom, derived from `top`; null without geometry. */
  yFraction: number | null;
  /** The trimmed target line text (capped), for tooltips / previews. */
  snippet: string;
}

/** One clickable mention of a target inside a page's text. */
export interface PdfReferenceMention {
  /** The target key this mention resolves to (always present in `targets`). */
  key: string;
  /** Index into the input pages array. */
  pageIndex: number;
  /** FOLDED-text range (same space as find matches → reuse rect measurement). */
  start: number;
  /** Exclusive folded end offset. */
  end: number;
  /** The mention's visible raw text, e.g. "式(1.2)" / "Eq. (1.2)". */
  label: string;
}

/** The full document reference index. */
export interface PdfReferenceIndex {
  /** targetKey → target location. */
  targets: ReadonlyMap<string, PdfReferenceTarget>;
  /** pageIndex → mentions with an existing target, ordered by offset. */
  mentionsByPage: ReadonlyMap<number, PdfReferenceMention[]>;
  /**
   * The searchable entry built per page (1:1 with the inputs; null where the
   * input was null/failed) — the offset maps rect measurement needs.
   */
  entries: ReadonlyArray<PdfPageSearchText | null>;
}

/** Per-page scan budget so one pathological text stream can't stall the UI. */
const MAX_SCAN_CHARS = 60_000;
/** Cap on linked mentions per page (dense bibliographies stay bounded). */
const MAX_MENTIONS_PER_PAGE = 300;
/** Snippet length cap for target previews. */
const MAX_SNIPPET_CHARS = 140;

/**
 * A reference label: optional 1–2 letter prefix ("A.3"), dotted/hyphenated
 * number groups ("1.2.3", "2-3"), optional letter suffix ("3a"), optional
 * prime. Shared by targets and mentions so both sides key identically.
 */
const LABEL_SRC = String.raw`(?:[A-Za-z]{1,2}[.\-])?\d{1,4}(?:[.\-]\d{1,3})*[A-Za-z]?['′’]?`;

/**
 * Canonicalize a matched label for keying: fullwidth → ASCII (NFKC), spaces
 * stripped, primes unified, letter PREFIX uppercased ("a.3" → "A.3") and
 * letter SUFFIX lowercased ("3A" → "3a") so casing conventions on the two
 * sides of a reference still meet on one key.
 */
function canonicalLabel(raw: string): string {
  let label = raw.normalize("NFKC").replace(/\s+/g, "").replace(/[′’]/g, "'");
  label = label.replace(
    /^([A-Za-z]{1,2})([.\-])/,
    (_m, prefix: string, sep: string) => `${prefix.toUpperCase()}${sep}`,
  );
  label = label.replace(
    /(\d)([A-Za-z])(?='?$)/,
    (_m, digit: string, suffix: string) => `${digit}${suffix.toLowerCase()}`,
  );
  return label;
}

// ---- Target detection (per raw-text line) ---------------------------------

interface TargetCandidate {
  key: string;
  /** Raw-text range of the MARKER (the "Figure 3:" / "(1.2)" phrase itself). */
  markerStart: number;
  markerEnd: number;
  snippet: string;
}

/**
 * Equation number at the END of a line: `… (1.2)` / whole-line `(1.2)`, with
 * optional trailing punctuation. Adapted from `isEquationNumberLine` in
 * text-boxes.ts (the formula pipeline's equation-number recognizer), narrowed
 * to the paren form — bracket-form labels collide with `[k]` citations.
 */
function eqTargetRe(): RegExp {
  return new RegExp(
    String.raw`([(（]\s*(${LABEL_SRC})\s*[)）])\s*[.,:;。，；]?$`,
  );
}

/** En figure/table caption at line start: "Figure 3:", "Table 2.1 Results". */
function enCaptionTargetRe(): RegExp {
  return new RegExp(
    String.raw`^(figure|fig\.?|table|tab\.?)\s+(${LABEL_SRC})\s*(?=$|[:.：。．]|\s+[A-Z一-鿿])`,
    "i",
  );
}

/** zh figure/table caption prefix at line start: "图 2.8 …" / "表 3：…". */
function zhCaptionTargetRe(): RegExp {
  return new RegExp(String.raw`^(图|表)\s*(${LABEL_SRC})`);
}

/** En theorem-like heading: "Theorem 3.1." / "Lemma 2 (Name)". */
function enTheoremTargetRe(): RegExp {
  return new RegExp(
    String.raw`^(theorem|lemma|proposition|corollary|definition|example|remark|algorithm)\s+(${LABEL_SRC})\s*(?=$|[.:(（．。：])`,
    "i",
  );
}

/** zh theorem-like heading prefix: "定理 3.1（…）" / "引理 2." */
function zhTheoremTargetRe(): RegExp {
  return new RegExp(
    String.raw`^(定理|引理|命题|推论|定义|例|注|算法)\s*(${LABEL_SRC})`,
  );
}

/** Bibliography entry: "[12] A. Author, …" (marker + entry text). */
function bibTargetRe(): RegExp {
  return /^(\[(\d{1,3})\])\s*\S/;
}

/** Multi-level numbered section heading: "2.3 Convergence analysis". */
function secTargetRe(): RegExp {
  return /^(\d{1,2}(?:\.\d{1,3})+)\.?\s+[^\s\d]/;
}

/** Chapter headings: "Chapter 3" / "第 3 章"; zh section "第 3 节". */
function enChapterTargetRe(): RegExp {
  return /^chapter\s+(\d{1,3})(?![\d.])/i;
}
function zhChapterTargetRe(): RegExp {
  return /^第\s*(\d{1,3})\s*(章|节)/;
}

const EN_THM_KINDS: Readonly<Record<string, string>> = {
  theorem: "thm",
  theorems: "thm",
  lemma: "lem",
  lemmas: "lem",
  proposition: "prop",
  propositions: "prop",
  corollary: "cor",
  corollaries: "cor",
  definition: "def",
  definitions: "def",
  example: "ex",
  examples: "ex",
  remark: "rem",
  remarks: "rem",
  algorithm: "alg",
  algorithms: "alg",
};

const ZH_THM_KINDS: Readonly<Record<string, string>> = {
  定理: "thm",
  引理: "lem",
  命题: "prop",
  推论: "cor",
  定义: "def",
  例: "ex",
  注: "rem",
  算法: "alg",
};

/**
 * Characters that, immediately after a zh line-start keyword+number, signal a
 * line-wrapped RUNNING-TEXT reference rather than a caption/heading ("…如\n
 * 图 3 所示…"). Precision over recall: such lines produce no target.
 */
const ZH_TARGET_CONTINUATION = new Set([
  "所",
  "中",
  "可",
  "示",
  "与",
  "及",
  "和",
  "之",
  "的",
]);

/** Whether the zh caption/heading match is followed by caption-like content. */
function zhTargetTailAccepted(trimmedLine: string, matchEnd: number): boolean {
  const tail = trimmedLine.slice(matchEnd).replace(/^[\s]+/, "");
  if (tail.length === 0) return true;
  const first = tail[0];
  if (ZH_TARGET_CONTINUATION.has(first)) return false;
  // A digit right after means the label regex under-consumed — reject; a
  // lowercase latin run reads like wrapped prose, not a caption.
  if (/[0-9a-z]/.test(first)) return false;
  return true;
}

/** A raw-offset mention candidate produced by the line scan (TOC rows). */
interface RawLineMention {
  key: string;
  rawStart: number;
  rawEnd: number;
}

interface PageLineScan {
  targets: TargetCandidate[];
  /** TOC-row mentions: the heading text of a leader-dots row links forward. */
  tocMentions: RawLineMention[];
}

/** Scan one page's raw text for target candidates, in line order. */
function scanTargetsInPage(rawText: string): PageLineScan {
  const out: TargetCandidate[] = [];
  const tocMentions: RawLineMention[] = [];
  const scanText =
    rawText.length > MAX_SCAN_CHARS ? rawText.slice(0, MAX_SCAN_CHARS) : rawText;
  let offset = 0;
  for (const line of scanText.split("\n")) {
    const lineOffset = offset;
    offset += line.length + 1;
    const leading = line.length - line.trimStart().length;
    const trimmed = line.trim();
    if (!trimmed) continue;
    // Table-of-contents / list-of-figures rows ("2.3 Analysis ...... 45",
    // leader-dots pattern shared with isDocumentIndexLine in text-boxes.ts)
    // must never become targets — first-occurrence-wins would send every
    // section/figure jump to the TOC page. Instead a section/chapter TOC row
    // becomes a MENTION over its heading text, clickable toward the real
    // heading once one exists.
    const leader = /(?:\.\s*){3,}\d{1,4}\s*$/.exec(trimmed);
    if (leader) {
      const headText = trimmed.slice(0, leader.index).trimEnd();
      let key: string | null = null;
      const secRow = /^(\d{1,2}(?:\.\d{1,3})+)[.)]?\s+\S/.exec(headText);
      const enChapRow = /^chapter\s+(\d{1,3})(?![\d.])/i.exec(headText);
      const zhChapRow = /^第\s*(\d{1,3})\s*(章|节)/.exec(headText);
      if (secRow) key = `sec:${secRow[1]}`;
      else if (enChapRow) key = `chap:${enChapRow[1]}`;
      else if (zhChapRow) {
        key = `${zhChapRow[2] === "章" ? "chap" : "zsec"}:${zhChapRow[1]}`;
      }
      if (key && headText.length > 0) {
        tocMentions.push({
          key,
          rawStart: lineOffset + leading,
          rawEnd: lineOffset + leading + headText.length,
        });
      }
      continue;
    }
    const base = lineOffset + leading;
    const snippet = trimmed.slice(0, MAX_SNIPPET_CHARS);
    const push = (key: string, markerStart: number, markerEnd: number): void => {
      out.push({ key, markerStart, markerEnd, snippet });
    };

    // Bibliography entry — checked before the equation pattern so a "[12] …"
    // entry line is never misread (eq targets are paren-form only anyway).
    const bib = bibTargetRe().exec(trimmed);
    if (bib) {
      push(`bib:${bib[2]}`, base + bib.index, base + bib.index + bib[1].length);
    }

    // Equation number at line end.
    const eq = eqTargetRe().exec(trimmed);
    if (eq && (eq.index === 0 || /\s/.test(trimmed[eq.index - 1] ?? ""))) {
      push(
        `eq:${canonicalLabel(eq[2])}`,
        base + eq.index,
        base + eq.index + eq[1].length,
      );
    }

    // Figure / table captions.
    const enCap = enCaptionTargetRe().exec(trimmed);
    if (enCap) {
      const kind = enCap[1].toLowerCase().startsWith("f") ? "fig" : "tab";
      push(
        `${kind}:${canonicalLabel(enCap[2])}`,
        base + enCap.index,
        base + enCap.index + enCap[0].length,
      );
    }
    const zhCap = zhCaptionTargetRe().exec(trimmed);
    if (zhCap && zhTargetTailAccepted(trimmed, zhCap[0].length)) {
      const kind = zhCap[1] === "图" ? "fig" : "tab";
      push(
        `${kind}:${canonicalLabel(zhCap[2])}`,
        base + zhCap.index,
        base + zhCap.index + zhCap[0].length,
      );
    }

    // Theorem-like headings.
    const enThm = enTheoremTargetRe().exec(trimmed);
    if (enThm) {
      const kind = EN_THM_KINDS[enThm[1].toLowerCase()];
      if (kind) {
        push(
          `${kind}:${canonicalLabel(enThm[2])}`,
          base + enThm.index,
          base + enThm.index + enThm[0].length,
        );
      }
    }
    const zhThm = zhTheoremTargetRe().exec(trimmed);
    if (zhThm && zhTargetTailAccepted(trimmed, zhThm[0].length)) {
      const kind = ZH_THM_KINDS[zhThm[1]];
      if (kind) {
        push(
          `${kind}:${canonicalLabel(zhThm[2])}`,
          base + zhThm.index,
          base + zhThm.index + zhThm[0].length,
        );
      }
    }

    // Numbered section / chapter headings.
    const sec = secTargetRe().exec(trimmed);
    if (sec) {
      push(`sec:${sec[1]}`, base, base + sec[1].length);
    }
    const enChap = enChapterTargetRe().exec(trimmed);
    if (enChap) {
      push(`chap:${enChap[1]}`, base, base + enChap[0].length);
    }
    const zhChap = zhChapterTargetRe().exec(trimmed);
    if (zhChap) {
      const kind = zhChap[2] === "章" ? "chap" : "zsec";
      push(`${kind}:${zhChap[1]}`, base, base + zhChap[0].length);
    }
  }
  return { targets: out, tocMentions };
}

// ---- Mention detection (across each page's raw text) -----------------------

interface RawMention {
  key: string;
  rawStart: number;
  rawEnd: number;
}

interface MentionSpec {
  re: () => RegExp;
  /** Resolve a match to a target key, or null to reject it. */
  keyOf: (m: RegExpExecArray, rawText: string) => string | null;
}

/** True when the char immediately before `index` is a letter/digit (so a
 * latin keyword match inside a longer word — "prefigure 3" — is rejected). */
function precededByWordChar(rawText: string, index: number): boolean {
  if (index <= 0) return false;
  const before = rawText[index - 1];
  return /[\p{L}\p{N}]/u.test(before);
}

/** True when the char before `index` is in the given zh false-positive set
 * ("方式(1)" must not read as 式(1), "代表 3" is not 表 3). */
function precededByZhBlock(
  rawText: string,
  index: number,
  block: string,
): boolean {
  if (index <= 0) return false;
  return block.includes(rawText[index - 1]);
}

const ZH_EQ_PRE_BLOCK = "方形模格样各等算招板纸架图程公款";
const ZH_TABLE_PRE_BLOCK = "代列报课时手仪钟图量额";
const ZH_EXAMPLE_PRE_BLOCK = "比案条实病特举示图范样";
const ZH_NOTE_PRE_BLOCK = "备关标脚附";

const MENTION_SPECS: ReadonlyArray<MentionSpec> = [
  // Equations, paren form (zh + en): 式(1.2) / 公式（3） / Eq. (1.2).
  {
    re: () =>
      new RegExp(
        String.raw`(式|公式|方程|[Ee]quations?|[Ee][Qq][Ss]?\.?)\s*[(（]\s*(${LABEL_SRC})\s*[)）]`,
        "g",
      ),
    keyOf: (m, rawText) => {
      const keyword = m[1];
      if (/^[A-Za-z]/.test(keyword)) {
        if (precededByWordChar(rawText, m.index)) return null;
      } else if (
        keyword === "式" &&
        precededByZhBlock(rawText, m.index, ZH_EQ_PRE_BLOCK)
      ) {
        return null;
      }
      return `eq:${canonicalLabel(m[2])}`;
    },
  },
  // Equations, bare-number form for unambiguous multi-char keywords:
  // 公式 3.1 / Equation 5 (single-char 式 REQUIRES parens, above).
  {
    re: () =>
      new RegExp(
        String.raw`(公式|方程|[Ee]quations?|[Ee][Qq][Ss]?\.)\s*(${LABEL_SRC})`,
        "g",
      ),
    keyOf: (m, rawText) => {
      if (/^[A-Za-z]/.test(m[1]) && precededByWordChar(rawText, m.index)) {
        return null;
      }
      return `eq:${canonicalLabel(m[2])}`;
    },
  },
  // Figures: 图 2.8 / Fig. 3 / Figure 3a.
  {
    re: () =>
      new RegExp(String.raw`(图|[Ff]igs?\.?|[Ff]igures?)\s*(${LABEL_SRC})`, "g"),
    keyOf: (m, rawText) => {
      if (/^[A-Za-z]/.test(m[1]) && precededByWordChar(rawText, m.index)) {
        return null;
      }
      return `fig:${canonicalLabel(m[2])}`;
    },
  },
  // Tables: 表 3 / Table 2.1 / Tab. 4.
  {
    re: () =>
      new RegExp(String.raw`(表|[Tt]ab\.?|[Tt]ables?)\s*(${LABEL_SRC})`, "g"),
    keyOf: (m, rawText) => {
      const keyword = m[1];
      if (/^[A-Za-z]/.test(keyword)) {
        if (precededByWordChar(rawText, m.index)) return null;
      } else if (precededByZhBlock(rawText, m.index, ZH_TABLE_PRE_BLOCK)) {
        return null;
      }
      return `tab:${canonicalLabel(m[2])}`;
    },
  },
  // Theorem-like (en): Theorem 3.1 / Lemma 2 / Algorithm 1 …
  {
    re: () =>
      new RegExp(
        String.raw`([Tt]heorems?|[Ll]emmas?|[Pp]ropositions?|[Cc]orollar(?:y|ies)|[Dd]efinitions?|[Ee]xamples?|[Rr]emarks?|[Aa]lgorithms?)\s+(${LABEL_SRC})`,
        "g",
      ),
    keyOf: (m, rawText) => {
      if (precededByWordChar(rawText, m.index)) return null;
      const kind = EN_THM_KINDS[m[1].toLowerCase()];
      return kind ? `${kind}:${canonicalLabel(m[2])}` : null;
    },
  },
  // Theorem-like (zh): 定理 3.1 / 引理 2 / 例 3 …
  {
    re: () =>
      new RegExp(
        String.raw`(定理|引理|命题|推论|定义|例|注|算法)\s*(${LABEL_SRC})`,
        "g",
      ),
    keyOf: (m, rawText) => {
      const keyword = m[1];
      if (
        keyword === "例" &&
        precededByZhBlock(rawText, m.index, ZH_EXAMPLE_PRE_BLOCK)
      ) {
        return null;
      }
      if (
        keyword === "注" &&
        precededByZhBlock(rawText, m.index, ZH_NOTE_PRE_BLOCK)
      ) {
        return null;
      }
      const kind = ZH_THM_KINDS[keyword];
      return kind ? `${kind}:${canonicalLabel(m[2])}` : null;
    },
  },
  // Sections: Section 2.3 / Sec. 4.1 / §3.2 / 2.3 节 / 第 3 章 / 第 2 节.
  {
    re: () =>
      new RegExp(
        String.raw`([Ss]ections?|[Ss]ect?\.|§)\s*(${LABEL_SRC})`,
        "g",
      ),
    keyOf: (m, rawText) => {
      if (/^[A-Za-z]/.test(m[1]) && precededByWordChar(rawText, m.index)) {
        return null;
      }
      return `sec:${canonicalLabel(m[2])}`;
    },
  },
  {
    re: () => /(第)\s*(\d{1,3})\s*(章|节)/g,
    keyOf: (m) => (m[3] === "章" ? `chap:${m[2]}` : `zsec:${m[2]}`),
  },
  {
    re: () => /()(\d{1,2}(?:\.\d{1,3})+)\s*节/g,
    keyOf: (m, rawText) => {
      // "第 2.3 节" is handled here too; avoid double count with the plain
      // 第N节 spec via overlap dedupe. Reject digits glued to the left.
      if (precededByWordChar(rawText, m.index)) return null;
      return `sec:${m[2]}`;
    },
  },
  // Bibliography citations: [12].
  {
    re: () => /(\[)(\d{1,3})\]/g,
    keyOf: (m) => `bib:${m[2]}`,
  },
];

/** Scan one page's raw text for mention candidates (unfiltered). */
function scanMentionsInPage(rawText: string): RawMention[] {
  const scanText =
    rawText.length > MAX_SCAN_CHARS ? rawText.slice(0, MAX_SCAN_CHARS) : rawText;
  const out: RawMention[] = [];
  for (const spec of MENTION_SPECS) {
    const re = spec.re();
    let m = re.exec(scanText);
    while (m) {
      const key = spec.keyOf(m, scanText);
      if (key && m[0].length > 0) {
        out.push({ key, rawStart: m.index, rawEnd: m.index + m[0].length });
      }
      if (re.lastIndex <= m.index) re.lastIndex = m.index + 1;
      m = re.exec(scanText);
    }
  }
  return out;
}

// ---- Offset mapping ---------------------------------------------------------

/**
 * Map a RAW-text range to the FOLDED range that covers it (the inverse of
 * `rawRangeForNormalizedRange`, via binary search over the monotonic
 * normStarts/normEnds arrays). Null when the raw range folded away entirely.
 */
function foldedRangeForRawRange(
  entry: Pick<PdfPageSearchText, "normStarts" | "normEnds">,
  rawStart: number,
  rawEnd: number,
): { start: number; end: number } | null {
  const n = entry.normStarts.length;
  if (n === 0 || rawEnd <= rawStart) return null;
  // First folded char whose source extends past rawStart.
  let lo = 0;
  let hi = n;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if ((entry.normEnds[mid] ?? 0) > rawStart) hi = mid;
    else lo = mid + 1;
  }
  const start = lo;
  // One past the last folded char whose source begins before rawEnd.
  lo = 0;
  hi = n;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if ((entry.normStarts[mid] ?? 0) < rawEnd) lo = mid + 1;
    else hi = mid;
  }
  const end = lo;
  if (end <= start) return null;
  return { start, end };
}

/** PDF user-space top of the first geometric item overlapping a raw range. */
function markerTop(
  items: ReadonlyArray<PdfSearchTextItem>,
  entry: PdfPageSearchText,
  rawStart: number,
  rawEnd: number,
): number | null {
  const indexes = itemsForRange(
    entry.itemStarts,
    rawStart,
    rawEnd,
    entry.rawText.length,
  );
  for (const idx of indexes) {
    const item = items[idx];
    const t = item?.transform;
    if (t && t.length >= 6 && Number.isFinite(t[5])) {
      const height =
        typeof item.height === "number" && Number.isFinite(item.height)
          ? Math.max(0, item.height)
          : 0;
      return (t[5] as number) + height;
    }
  }
  return null;
}

// ---- Index assembly ---------------------------------------------------------

/**
 * Build the document reference index from per-page inputs (index = the
 * viewer's visible-page order, matching the find pipeline's pageIndex).
 * Null inputs (failed extraction) contribute nothing but keep positions.
 */
export function buildReferenceIndex(
  pages: ReadonlyArray<PdfReferencePageInput | null>,
): PdfReferenceIndex {
  const entries: Array<PdfPageSearchText | null> = [];
  const itemsByPage: Array<PdfSearchTextItem[] | null> = [];
  for (const page of pages) {
    if (!page || !Array.isArray(page.items)) {
      entries.push(null);
      itemsByPage.push(null);
      continue;
    }
    const items: PdfSearchTextItem[] = page.items.map((item) =>
      typeof item === "string" ? { str: item } : item,
    );
    itemsByPage.push(items);
    entries.push(buildSearchablePageText(items));
  }

  // Pass 1: targets (first occurrence in document order wins). Marker ranges
  // of ALL candidates (winners and duplicates) are kept so a mention never
  // links a caption/heading/entry to itself.
  const targets = new Map<string, PdfReferenceTarget>();
  const markersByPage = new Map<number, Array<{ start: number; end: number }>>();
  const tocMentionsByPage = new Map<number, RawLineMention[]>();
  for (let pageIndex = 0; pageIndex < entries.length; pageIndex += 1) {
    const entry = entries[pageIndex];
    const items = itemsByPage[pageIndex];
    if (!entry || !items || !entry.rawText.trim()) continue;
    const scan = scanTargetsInPage(entry.rawText);
    if (scan.tocMentions.length > 0) {
      tocMentionsByPage.set(pageIndex, scan.tocMentions);
    }
    const candidates = scan.targets;
    if (candidates.length === 0) continue;
    const markers = markersByPage.get(pageIndex) ?? [];
    for (const candidate of candidates) {
      markers.push({ start: candidate.markerStart, end: candidate.markerEnd });
      if (targets.has(candidate.key)) continue;
      const top = markerTop(
        items,
        entry,
        candidate.markerStart,
        candidate.markerEnd,
      );
      const rawHeight = pages[pageIndex]?.pageHeight;
      const pageHeight =
        typeof rawHeight === "number" &&
        Number.isFinite(rawHeight) &&
        rawHeight > 0
          ? rawHeight
          : null;
      const yFraction =
        top != null && pageHeight != null
          ? Math.max(0, Math.min(1, (pageHeight - top) / pageHeight))
          : null;
      targets.set(candidate.key, {
        key: candidate.key,
        pageIndex,
        top,
        yFraction,
        snippet: candidate.snippet,
      });
    }
    if (markers.length > 0) markersByPage.set(pageIndex, markers);
  }

  // Pass 2: mentions — keep only those with an existing target that are not a
  // target marker themselves, dedupe overlaps, map to folded offsets.
  const mentionsByPage = new Map<number, PdfReferenceMention[]>();
  if (targets.size > 0) {
    for (let pageIndex = 0; pageIndex < entries.length; pageIndex += 1) {
      const entry = entries[pageIndex];
      if (!entry || !entry.rawText.trim()) continue;
      const markers = markersByPage.get(pageIndex) ?? [];
      const raw = scanMentionsInPage(entry.rawText)
        .concat(tocMentionsByPage.get(pageIndex) ?? [])
        .filter((mention) => targets.has(mention.key))
        .filter(
          (mention) =>
            !markers.some(
              (marker) =>
                mention.rawStart < marker.end && mention.rawEnd > marker.start,
            ),
        )
        .sort((a, b) => a.rawStart - b.rawStart || b.rawEnd - a.rawEnd);
      const mentions: PdfReferenceMention[] = [];
      let lastEnd = -1;
      for (const mention of raw) {
        if (mentions.length >= MAX_MENTIONS_PER_PAGE) break;
        if (mention.rawStart < lastEnd) continue;
        const folded = foldedRangeForRawRange(
          entry,
          mention.rawStart,
          mention.rawEnd,
        );
        if (!folded) continue;
        lastEnd = mention.rawEnd;
        mentions.push({
          key: mention.key,
          pageIndex,
          start: folded.start,
          end: folded.end,
          label: entry.rawText.slice(mention.rawStart, mention.rawEnd),
        });
      }
      if (mentions.length > 0) mentionsByPage.set(pageIndex, mentions);
    }
  }

  return { targets, mentionsByPage, entries };
}
