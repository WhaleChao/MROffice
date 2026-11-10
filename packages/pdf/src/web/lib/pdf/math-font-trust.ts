/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Per-font text-extraction trust for PDF math fonts.
 *
 * Office-exported PDFs (Word/PowerPoint equations) embed math-font subsets
 * (e.g. `AAAABM+CambriaMath`) whose ToUnicode CMaps are systematically
 * incomplete: mapped charcodes decode correctly, while unmapped ones fall back
 * to the raw charcode in pdf.js, so `x→c` extracts as `2→3`. The rendered
 * canvas is correct — only the extracted *text* is garbage — and a wrong
 * character stream silently becomes a wrong-but-plausible LaTeX seed.
 *
 * This module parses the document's font dictionaries once and reports, per
 * embedded font, the set of characters the ToUnicode CMap can actually
 * produce. Extraction then treats math-font runs containing characters
 * outside that set as untrusted: they still group into formula boxes (the
 * font name itself is a strong formula signal) but must not fabricate LaTeX.
 *
 * Keyed by the RAW BaseFont name (subset tag included): each subset has its
 * own charcode assignment, so merging `AAAABI+CambriaMath` with
 * `AAAABM+CambriaMath` would let one subset's real mapping vouch for another
 * subset's fallback garbage.
 */

import { loadPdfForInspection } from "./encrypted-guard";
import {
  PDFDict,
  PDFDocument,
  PDFName,
  PDFRawStream,
  decodePDFRawStream,
} from "pdf-lib";

import { stripPdfFontSubsetTag } from "./font-resolve";

export interface PdfFontTextTrust {
  /** Raw BaseFont name, subset tag included (e.g. `AAAABM+CambriaMath`). */
  rawFontName: string;
  /** Characters the font's ToUnicode CMap can produce (code-point strings). */
  trustedChars: ReadonlySet<string>;
  /** Number of charcodes with a ToUnicode entry. */
  mappedCharcodeCount: number;
  /** Number of charcodes the font declares glyph widths for (0 if unknown). */
  declaredCharcodeCount: number;
}

export interface PdfTextTrustReport {
  /** Keyed by raw BaseFont name including the subset tag. */
  fonts: ReadonlyMap<string, PdfFontTextTrust>;
}

export type PdfTextTrustVerdict = "trusted" | "untrusted" | "unknown";

const MATH_SOURCE_FONT_RE =
  /(?:^|[-+_ .,])math(?:$|[-+_ .,])|cambria ?math|latin ?modern ?math|lmmath|newcm(?:math)?|stix|xits|asana|fira ?math|tex ?gyre.*math|euclid(?: ?(?:math|symbol|extra))?|mathematica|cmmi|cmsy|cmex|cmbsy|cmmib|msam|msbm|eufm|eusm|eusb|mt(?:mi|sy|ex)\b|mathpack/i;

/**
 * True when a PDF source font name looks like a dedicated math face. Accepts
 * both raw (`AAAABM+CambriaMath`) and normalized (`CambriaMath`) names.
 */
export function isMathSourceFontName(name: string | null | undefined): boolean {
  const stripped = stripPdfFontSubsetTag(name);
  if (!stripped) return false;
  return MATH_SOURCE_FONT_RE.test(stripped);
}

function utf16beToString(hexDigits: string): string {
  const units: number[] = [];
  for (let i = 0; i + 4 <= hexDigits.length; i += 4) {
    const unit = Number.parseInt(hexDigits.slice(i, i + 4), 16);
    if (Number.isNaN(unit)) return "";
    units.push(unit);
  }
  // Odd trailing bytes (malformed CMaps) are ignored; a 2-digit tail such as
  // <41> is a single byte value some producers emit for ASCII.
  if (units.length === 0 && hexDigits.length === 2) {
    const byte = Number.parseInt(hexDigits, 16);
    if (!Number.isNaN(byte)) units.push(byte);
  }
  return String.fromCharCode(...units);
}

/**
 * Parse the bfchar/bfrange sections of a ToUnicode CMap program. Returns a
 * charcode → mapped-string map. Pure text parsing so it is unit-testable
 * without pdf-lib; tolerates the array form of bfrange.
 */
export function parseToUnicodeCMap(cmapText: string): Map<number, string> {
  const map = new Map<number, string>();

  const charSections = cmapText.matchAll(/beginbfchar([\s\S]*?)endbfchar/g);
  for (const section of charSections) {
    const pairs = section[1].matchAll(/<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>/g);
    for (const pair of pairs) {
      const code = Number.parseInt(pair[1], 16);
      const value = utf16beToString(pair[2]);
      if (Number.isFinite(code) && value) map.set(code, value);
    }
  }

  const rangeSections = cmapText.matchAll(/beginbfrange([\s\S]*?)endbfrange/g);
  for (const section of rangeSections) {
    const body = section[1];
    const rangeRe =
      /<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>\s*(?:<([0-9a-fA-F]+)>|\[((?:\s*<[0-9a-fA-F]+>)+)\s*\])/g;
    for (const range of body.matchAll(rangeRe)) {
      const low = Number.parseInt(range[1], 16);
      const high = Number.parseInt(range[2], 16);
      if (!Number.isFinite(low) || !Number.isFinite(high) || high < low) {
        continue;
      }
      if (high - low > 0xffff) continue; // defensive: malformed huge range
      if (range[3] != null) {
        const base = utf16beToString(range[3]);
        const baseCp = base.codePointAt(0);
        if (base === "" || baseCp == null) continue;
        const suffix = base.slice(String.fromCodePoint(baseCp).length);
        for (let code = low; code <= high; code++) {
          map.set(
            code,
            String.fromCodePoint(baseCp + (code - low)) + suffix,
          );
        }
      } else if (range[4] != null) {
        const values = [...range[4].matchAll(/<([0-9a-fA-F]+)>/g)].map((m) =>
          utf16beToString(m[1]),
        );
        for (let i = 0; i < values.length && low + i <= high; i++) {
          if (values[i]) map.set(low + i, values[i]);
        }
      }
    }
  }
  return map;
}

function decodeStreamText(
  dict: PDFDict,
  ref: unknown,
): string | null {
  try {
    const stream = dict.context.lookup(ref as Parameters<typeof dict.context.lookup>[0]);
    if (!(stream instanceof PDFRawStream)) return null;
    const bytes = decodePDFRawStream(stream).decode();
    let out = "";
    for (const byte of bytes) out += String.fromCharCode(byte);
    return out;
  } catch {
    return null;
  }
}

function fontTrustFromDict(dict: PDFDict): PdfFontTextTrust | null {
  const baseFontValue = dict.get(PDFName.of("BaseFont"));
  if (!baseFontValue) return null;
  const rawFontName = String(baseFontValue).replace(/^\//, "");
  if (!rawFontName) return null;

  let mapped: Map<number, string> | null = null;
  const toUnicodeRef = dict.get(PDFName.of("ToUnicode"));
  if (toUnicodeRef) {
    const text = decodeStreamText(dict, toUnicodeRef);
    if (text != null) mapped = parseToUnicodeCMap(text);
  }

  let declaredCharcodeCount = 0;
  try {
    const widths = dict.context.lookup(dict.get(PDFName.of("Widths")));
    const size = (widths as { size?: () => number } | undefined)?.size?.();
    if (typeof size === "number" && Number.isFinite(size)) {
      declaredCharcodeCount = size;
    }
  } catch {
    // Widths are advisory for coverage stats only.
  }

  const trustedChars = new Set<string>();
  if (mapped) {
    for (const value of mapped.values()) {
      for (const char of value) trustedChars.add(char);
    }
  }
  return {
    rawFontName,
    trustedChars,
    mappedCharcodeCount: mapped?.size ?? 0,
    declaredCharcodeCount,
  };
}

/**
 * Parse a PDF once and report per-font extraction trust. Only fonts that
 * look like math faces are analyzed — prose fonts with imperfect CMaps show
 * ordinary mojibake, which is out of scope here. Never throws: an unparseable
 * document yields an empty report (all verdicts "unknown" → no behavior
 * change).
 */
export async function analyzePdfTextTrust(
  pdfBytes: Uint8Array | ArrayBuffer,
): Promise<PdfTextTrustReport> {
  const fonts = new Map<string, PdfFontTextTrust>();
  try {
    const doc = await loadPdfForInspection(
      pdfBytes instanceof Uint8Array ? pdfBytes : new Uint8Array(pdfBytes),
      { updateMetadata: false },
    );
    // Encrypted bytes carry no readable font programs — no trust signal
    // either way; land in the same fallback as a parse failure.
    if (!doc) throw new Error("encrypted");
    for (const [, obj] of doc.context.enumerateIndirectObjects()) {
      if (!(obj instanceof PDFDict)) continue;
      if (String(obj.get(PDFName.of("Type"))) !== "/Font") continue;
      const subtype = String(obj.get(PDFName.of("Subtype")) || "");
      // Type0 wrappers carry BaseFont + ToUnicode at the top level too; the
      // descendant CIDFont has neither, so a plain dict walk covers both.
      if (subtype === "/Type3") continue;
      const trust = fontTrustFromDict(obj);
      if (!trust) continue;
      if (!isMathSourceFontName(trust.rawFontName)) continue;
      const existing = fonts.get(trust.rawFontName);
      if (existing) {
        // Same BaseFont name reused across font dicts (rare): merging unions
        // the trusted sets, which errs toward trusting — acceptable because
        // identical subset names imply identical charcode assignments.
        const union = new Set(existing.trustedChars);
        for (const char of trust.trustedChars) union.add(char);
        fonts.set(trust.rawFontName, {
          rawFontName: trust.rawFontName,
          trustedChars: union,
          mappedCharcodeCount: Math.max(
            existing.mappedCharcodeCount,
            trust.mappedCharcodeCount,
          ),
          declaredCharcodeCount: Math.max(
            existing.declaredCharcodeCount,
            trust.declaredCharcodeCount,
          ),
        });
      } else {
        fonts.set(trust.rawFontName, trust);
      }
    }
  } catch {
    return { fonts: new Map() };
  }
  return { fonts };
}

const TRUST_EXEMPT_CHAR_RE = /[\s ]/;

/**
 * Verdict for one extracted run.
 *
 * - "unknown": not a math font, or the document report has no entry for it —
 *   extraction behaves exactly as before.
 * - "trusted": math font and every non-whitespace character is producible by
 *   the font's ToUnicode CMap.
 * - "untrusted": math font with at least one character that can only be a
 *   raw-charcode fallback — the string content is unreliable.
 */
export function textTrustForSourceFont(
  report: PdfTextTrustReport | null | undefined,
  rawFontName: string | null | undefined,
  text: string,
): PdfTextTrustVerdict {
  if (!report || !rawFontName) return "unknown";
  const entry = report.fonts.get(rawFontName);
  if (!entry) return "unknown";
  for (const char of text) {
    if (TRUST_EXEMPT_CHAR_RE.test(char)) continue;
    if (!entry.trustedChars.has(char)) return "untrusted";
  }
  return "trusted";
}
