/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure document-metadata model for the PDF "Document Properties" editor
 * (Acrobat / WPS "File ▸ Properties"): Title / Author / Subject / Keywords /
 * document language.
 *
 * This module is the PURE half — it has ZERO pdf-lib / pdfjs / DOM / React
 * dependency so it can be exhaustively unit-tested in isolation and reused by
 * both the writer (lib/pdf/flatten-apply.ts, which calls pdf-lib's
 * setTitle/setAuthor/setSubject/setKeywords) and the dialog UI
 * (components/notebook/contents/pdf/PdfPropertiesDialog.tsx).
 *
 * Responsibilities:
 *  - Define the normalized {@link PdfMetadata} model (five string fields).
 *  - {@link normalizePdfMetadata}: trim, strip control characters (PDF info
 *    strings are single-line — newlines/tabs/NULs would corrupt the value or
 *    leak across the document-info dictionary), and cap each field to a sane
 *    length so a pathologically long paste can never bloat the output bytes.
 *  - {@link readPdfInfoMetadata}: map a pdfjs `getMetadata().info` object (a
 *    loose `Record<string, unknown>`) into the normalized model — the editor
 *    seeds its inputs from the PDF's CURRENT metadata.
 *  - {@link splitKeywords} / {@link joinKeywords}: keywords are stored & shown
 *    as one comma-separated string but pdf-lib's `setKeywords` takes an ARRAY;
 *    these bridge the two without losing user-entered separators.
 *  - {@link metadataIsEmpty} / {@link metadataChanged}: let the UI decide
 *    whether there is anything to write (so an untouched dialog never rewrites
 *    the info dictionary, and an empty field is faithfully cleared).
 *
 * Coordinate-free, render-free: this is just string hygiene + a tiny model.
 */

/**
 * Normalized PDF document metadata. Every field is a trimmed, control-char-free
 * single-line string; an EMPTY string means "no value" (and, on write, clears
 * any existing value). `keywords` is the human-facing comma-separated form;
 * {@link splitKeywords} produces the array pdf-lib's `setKeywords` expects.
 * `language` is either a canonical BCP-47 tag, an empty clear request, or an
 * invalid draft retained solely so the dialog can explain and correct it.
 */
export interface PdfMetadata {
  title: string;
  author: string;
  subject: string;
  keywords: string;
  /** Canonical BCP-47 tag written to the PDF catalog's `/Lang` entry. */
  language: string;
}

/**
 * Per-field maximum length (characters). PDF document-info strings are short by
 * convention; this cap is a generous backstop that prevents a multi-megabyte
 * paste from bloating the exported bytes (or the info dictionary) while never
 * truncating a realistic title/author/subject/keyword list. Applied AFTER
 * trimming + control-char removal so the cap counts only meaningful characters.
 */
export const PDF_METADATA_MAX_LENGTH = 2000;

/** The canonical empty model (all fields blank). */
export const EMPTY_PDF_METADATA: PdfMetadata = Object.freeze({
  title: "",
  author: "",
  subject: "",
  keywords: "",
  language: "",
});

/**
 * PDF 2.0 describes `/Lang` as a language identifier following BCP 47. Keep a
 * modest explicit limit: real tags are short, while an unbounded catalog value
 * would let a hostile source inflate both the properties form and every copy.
 */
export const PDF_LANGUAGE_MAX_LENGTH = 100;

const PDF_PRIVATE_USE_LANGUAGE_RE = /^x(?:-[a-z\d]{1,8})+$/i;









const PDF_GRANDFATHERED_LANGUAGE_TAGS: ReadonlyMap<string, string> = new Map(
  [
    "en-GB-oed", "i-ami", "i-bnn", "i-default", "i-enochian", "i-hak", "i-klingon", "i-lux",
    "i-mingo", "i-navajo", "i-pwn", "i-tao", "i-tay", "i-tsu", "sgn-BE-FR", "sgn-BE-NL",
    "sgn-CH-DE", "art-lojban", "cel-gaulish", "no-bok", "no-nyn", "zh-guoyu", "zh-hakka",
    "zh-min", "zh-min-nan", "zh-xiang",
  ].map((tag) => [tag.toLowerCase(), tag]),
);

/**
 * Normalize a PDF document-language tag without guessing from document text.
 * An empty value deliberately clears `/Lang`; `null` means malformed and must
 * be surfaced to the user rather than silently converted into a clear.
 *
 * `Intl.getCanonicalLocales` supplies the standards-aware casing and alias
 * rules for ordinary BCP-47 tags. Private-use-only tags are valid BCP 47 but
 * are rejected by some Intl implementations, so accept their closed ASCII
 * grammar separately and canonicalize them to lower case.
 */
export function normalizePdfLanguageTag(value: unknown): string | null {
  if (typeof value !== "string") return "";
  const candidate = value.replace(/\p{Cc}/gu, " ").trim();
  if (!candidate) return "";
  if (
    candidate.length > PDF_LANGUAGE_MAX_LENGTH ||
    /\s/.test(candidate) ||
    candidate.includes("_")
  ) {
    return null;
  }
  if (PDF_PRIVATE_USE_LANGUAGE_RE.test(candidate)) {
    return candidate.toLowerCase();
  }
  const grandfathered = PDF_GRANDFATHERED_LANGUAGE_TAGS.get(candidate.toLowerCase());
  if (grandfathered) return grandfathered;
  try {
    const canonical = Intl.getCanonicalLocales(candidate);
    return canonical.length === 1 ? canonical[0] : null;
  } catch {
    return null;
  }
}

/**
 * Sanitize a single metadata field value:
 *  1. coerce to string (a non-string input → ""),
 *  2. replace every CONTROL character (C0/C1, incl. newline / tab / NUL) with a
 *     single space so the value stays single-line and can never inject a raw
 *     newline into the PDF info dictionary,
 *  3. collapse runs of whitespace to one space and trim the ends,
 *  4. cap to {@link PDF_METADATA_MAX_LENGTH}.
 *
 * Exported for direct unit testing of the hygiene rules.
 */
export function sanitizeMetadataField(value: unknown): string {
  if (typeof value !== "string") {
    // pdfjs info values are usually strings, but Title can be absent or a
    // non-string (e.g. a boolean for IsAcroFormPresent leaking in) — coerce
    // only genuine strings; everything else is "no value".
    return "";
  }
  // Replace any Unicode control character (Cc) with a space, then collapse
  // whitespace. \p{Cc} covers C0 (0x00–0x1F, incl. \n \r \t) and C1
  // (0x7F–0x9F). The `u` flag is required for the property escape.
  const noControl = value.replace(/\p{Cc}/gu, " ");
  const collapsed = noControl.replace(/\s+/g, " ").trim();
  if (collapsed.length <= PDF_METADATA_MAX_LENGTH) return collapsed;
  // Trim again after slicing so a cut mid-trailing-space leaves no edge space.
  return collapsed.slice(0, PDF_METADATA_MAX_LENGTH).trim();
}

/**
 * Normalize a (possibly partial / dirty) metadata draft into the canonical
 * model. Missing fields default to "". Idempotent: normalizing an already
 * normalized model returns an equal model.
 */
export function normalizePdfMetadata(
  draft: Partial<PdfMetadata> | null | undefined,
): PdfMetadata {
  const d = draft ?? {};
  const languageDraft = sanitizeMetadataField(d.language);
  const canonicalLanguage = normalizePdfLanguageTag(languageDraft);
  return {
    title: sanitizeMetadataField(d.title),
    author: sanitizeMetadataField(d.author),
    subject: sanitizeMetadataField(d.subject),
    keywords: sanitizeMetadataField(d.keywords),
    // Preserve an invalid draft so the form can show and correct it. The UI and
    // writer both reject it; replacing it with "" here would fake a clear.
    language: canonicalLanguage ?? languageDraft,
  };
}

/**
 * Map a pdfjs `getMetadata().info` object into the normalized model. The info
 * dictionary uses Capitalized PDF keys (`Title`, `Author`, `Subject`,
 * `Keywords`); anything missing becomes "". Defensive against a non-object
 * input (returns the empty model).
 */
export function readPdfInfoMetadata(
  info: Record<string, unknown> | null | undefined,
): PdfMetadata {
  if (!info || typeof info !== "object") return { ...EMPTY_PDF_METADATA };
  return normalizePdfMetadata({
    title: typeof info.Title === "string" ? info.Title : "",
    author: typeof info.Author === "string" ? info.Author : "",
    subject: typeof info.Subject === "string" ? info.Subject : "",
    keywords: typeof info.Keywords === "string" ? info.Keywords : "",
    language: typeof info.Language === "string" ? info.Language : "",
  });
}

/** The pdf.js `Metadata` surface this module reads (XMP packet values). */
export interface PdfXmpMetadataLike {
  get(name: string): unknown;
}

function xmpText(value: unknown): string {
  if (typeof value === "string") return value;
  if (Array.isArray(value)) {
    // `dc:creator` is an ordered sequence of names; `dc:title` and
    // `dc:description` are language alternatives whose first entry pdf.js
    // returns for the default language.
    return value
      .filter((entry): entry is string => typeof entry === "string")
      .map((entry) => entry.trim())
      .filter(Boolean)
      .join(", ");
  }
  return "";
}

/**
 * Map a pdf.js `getMetadata().metadata` XMP packet into the normalized model:
 * `dc:title`, `dc:creator`, `dc:description`, `pdf:Keywords` and
 * `dc:language`. Anything missing becomes ""; a null / non-object input
 * returns the empty model.
 */
export function readPdfXmpMetadata(
  metadata: PdfXmpMetadataLike | null | undefined,
): PdfMetadata {
  if (!metadata || typeof metadata.get !== "function") {
    return { ...EMPTY_PDF_METADATA };
  }
  const read = (name: string): unknown => {
    try {
      return metadata.get(name);
    } catch {
      return null;
    }
  };
  return normalizePdfMetadata({
    title: xmpText(read("dc:title")),
    author: xmpText(read("dc:creator")),
    subject: xmpText(read("dc:description")),
    keywords: xmpText(read("pdf:Keywords")),
    language: xmpText(read("dc:language")).split(",")[0]?.trim() ?? "",
  });
}







export function mergePdfMetadataSources(
  xmp: PdfMetadata,
  info: PdfMetadata,
): PdfMetadata {
  return {
    title: xmp.title || info.title,
    author: xmp.author || info.author,
    subject: xmp.subject || info.subject,
    keywords: xmp.keywords || info.keywords,
    language: xmp.language || info.language,
  };
}

/**
 * Split a human-typed keyword string into a clean array for pdf-lib's
 * `setKeywords`. Accepts comma OR semicolon separators (both common), trims
 * each token, drops empties, and de-dupes (case-sensitively, preserving first
 * occurrence + order). An empty / whitespace input → `[]`.
 */
export function splitKeywords(keywords: string): string[] {
  if (typeof keywords !== "string") return [];
  const seen = new Set<string>();
  const out: string[] = [];
  for (const raw of keywords.split(/[,;]/)) {
    const token = raw.replace(/\s+/g, " ").trim();
    if (!token || seen.has(token)) continue;
    seen.add(token);
    out.push(token);
  }
  return out;
}

/**
 * Join a keyword array back into the canonical comma-separated display string
 * (inverse of {@link splitKeywords} for already-clean tokens). Used when
 * seeding the dialog from a source that exposes keywords as an array.
 */
export function joinKeywords(keywords: ReadonlyArray<string>): string {
  if (!Array.isArray(keywords)) return "";
  return splitKeywords(keywords.join(", ")).join(", ");
}

/** True when every field is empty (nothing to display / write). */
export function metadataIsEmpty(meta: PdfMetadata): boolean {
  const m = normalizePdfMetadata(meta);
  return !m.title && !m.author && !m.subject && !m.keywords && !m.language;
}

/**
 * True when `next` differs from `current` in ANY field (after normalizing
 * both). Lets the UI gate the write so an unedited dialog leaves the info
 * dictionary byte-for-byte untouched.
 */
export function metadataChanged(
  current: PdfMetadata,
  next: PdfMetadata,
): boolean {
  const a = normalizePdfMetadata(current);
  const b = normalizePdfMetadata(next);
  return (
    a.title !== b.title ||
    a.author !== b.author ||
    a.subject !== b.subject ||
    a.keywords !== b.keywords ||
    a.language !== b.language
  );
}
