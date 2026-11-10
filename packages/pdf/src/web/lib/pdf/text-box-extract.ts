/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  normalizePdfSourceFontName,
  pdfDisplayFontFamily,
  resolvePdfFontFamily,
} from "./font-resolve";
import {
  isMathSourceFontName,
  textTrustForSourceFont,
  type PdfTextTrustReport,
} from "./math-font-trust";
import {
  groupTextBoxes,
  type PdfTextItem,
  type PdfViewportLike,
  type TextBox,
} from "./text-boxes";

interface PdfTextContentItemLike {
  str?: unknown;
  transform?: unknown;
  width?: unknown;
  height?: unknown;
  fontName?: unknown;
  kind?: unknown;
  latex?: unknown;
}

export interface PdfTextContentLike {
  items?: unknown;
  styles?: Record<string, { fontFamily?: string }>;
}

export interface PdfTextContentPageLike {
  getTextContent: () => Promise<PdfTextContentLike>;
  streamTextContent?: () => ReadableStream<PdfTextContentLike>;
}

interface PdfCommonObjectsLike {
  has?: (id: string) => boolean;
  get?: (id: string) => unknown;
}

interface PdfFontSourceLike {
  commonObjs?: PdfCommonObjectsLike;
  /**
   * Per-document math-font extraction trust (see lib/pdf/math-font-trust.ts).
   * When provided, runs set in a math face whose characters cannot come from
   * the font's ToUnicode CMap are marked untrusted so no LaTeX seed is
   * fabricated from raw-charcode fallback text.
   */
  textTrust?: PdfTextTrustReport | null;
}

interface PdfFontObjectLike {
  name?: unknown;
  black?: unknown;
  bold?: unknown;
  italic?: unknown;
}

function textContentItems(textContent: PdfTextContentLike): unknown[] {
  return Array.isArray(textContent.items) ? textContent.items : [];
}

function compactText(text: string): string {
  return text.normalize("NFKC").replace(/\s+/g, "").trim();
}

function hasCjkText(text: string): boolean {
  return /[\u3400-\u9fff]/.test(text);
}

function isSubsequence(needle: string, haystack: string): boolean {
  if (needle.length === 0) return true;
  let index = 0;
  for (const char of haystack) {
    if (char === needle[index]) index += 1;
    if (index === needle.length) return true;
  }
  return false;
}

function shortCjkRepairCandidate(
  text: string,
  pageText: string,
): string | null {
  const compact = compactText(text);
  if (compact.length < 2 || compact.length > 12) return null;
  if (!hasCjkText(compact)) return null;

  const lines = pageText
    .split(/[\r\n]+/)
    .map((line) => line.replace(/\s+/g, " ").trim())
    .filter(Boolean);
  for (const line of lines) {
    const candidate = compactText(line);
    if (!hasCjkText(candidate)) continue;
    if (candidate.length <= compact.length) continue;
    if (candidate.length > compact.length + 2) continue;
    if (candidate.length > 12) continue;
    if (!isSubsequence(compact, candidate)) continue;
    return line;
  }
  return null;
}

export function repairShortCjkTextBoxesFromPageText(
  boxes: readonly TextBox[],
  pageText: string | undefined,
): TextBox[] {
  if (!pageText?.trim()) return [...boxes];
  return boxes.map((box) => {
    if (box.kind === "formula") return box;
    const repair = shortCjkRepairCandidate(box.text, pageText);
    if (!repair || repair === box.text) return box;
    const beforeLength = Math.max(1, compactText(box.text).length);
    const afterLength = Math.max(beforeLength, compactText(repair).length);
    return {
      ...box,
      text: repair,
      // The page-level hint recovered characters that were absent from the
      // glyph runs, so their old offsets no longer satisfy runs↔text.
      runs: undefined,
      width: Math.min(1 - box.x, box.width * (afterLength / beforeLength)),
    };
  });
}

async function textContentFromStream(
  stream: ReadableStream<PdfTextContentLike>,
): Promise<PdfTextContentLike> {
  const reader = stream.getReader();
  const items: unknown[] = [];
  const styles: Record<string, { fontFamily?: string }> = {};
  const readWithTimeout = async () => {
    let timeoutId: ReturnType<typeof setTimeout> | null = null;
    try {
      return await Promise.race([
        reader.read(),
        new Promise<never>((_, reject) => {
          timeoutId = setTimeout(
            () => reject(new Error("PDF text stream timed out")),
            1200,
          );
        }),
      ]);
    } finally {
      if (timeoutId !== null) clearTimeout(timeoutId);
    }
  };
  let finished = false;
  try {
    for (;;) {
      const { done, value } = await readWithTimeout();
      if (done) break;
      if (!value || typeof value !== "object") continue;
      items.push(...textContentItems(value));
      Object.assign(styles, value.styles ?? {});
    }
    finished = true;
  } finally {
    if (finished) {
      try {
        reader.releaseLock();
      } catch {
        // Released already.
      }
    } else {
      // Given up (a slow chunk, a bad one): the stream is read out in the
      // background and dropped, never cancelled. PDF.js keeps sending a
      // cancelled text stream's chunks and closes it after, and both throw
      // in its message handler (「Cannot enqueue a chunk into a closed
      // readable stream」, uncaught, on every slow page under load).
      void (async () => {
        try {
          for (;;) {
            const { done } = await reader.read();
            if (done) break;
          }
        } catch {
          // The stream failed: nothing of it is used.
        } finally {
          try {
            reader.releaseLock();
          } catch {
            // Released already.
          }
        }
      })();
    }
  }
  return { items, styles };
}

export async function readPdfTextContent(
  page: PdfTextContentPageLike,
): Promise<PdfTextContentLike> {
  if (typeof page.streamTextContent === "function") {
    try {
      const streamed = await textContentFromStream(page.streamTextContent());
      if (textContentItems(streamed).length > 0) return streamed;
    } catch {
      // Fall through to getTextContent: PDF.js streams are best-effort and can
      // fail on older browsers or malformed documents.
    }
  }
  return page.getTextContent();
}

function isFiniteNumber(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value);
}

function isTextTransform(value: unknown): value is number[] {
  return (
    Array.isArray(value) &&
    value.length >= 6 &&
    value.slice(0, 6).every(isFiniteNumber)
  );
}

function estimateItemWidth(
  str: string,
  transform: number[],
  height: unknown,
): number {
  const fontSize =
    (isFiniteNumber(height) && Math.abs(height)) ||
    Math.hypot(transform[2] ?? 0, transform[3] ?? 0) ||
    Math.hypot(transform[0] ?? 0, transform[1] ?? 0) ||
    1;
  const glyphCount = Array.from(str.trim() || str).length || 1;
  return Math.max(fontSize * 0.35, glyphCount * fontSize * 0.52);
}

function shouldUseEstimatedWidth(
  str: string,
  reportedWidth: number,
  estimatedWidth: number,
): boolean {
  const glyphCount = Array.from(str.trim()).length;
  if (glyphCount < 4) return false;
  if (estimatedWidth <= 0) return false;
  if (reportedWidth >= estimatedWidth * 0.45) return false;
  // Prefer the estimate only for prose-like chunks. Short math atoms and page
  // labels can legitimately have tiny advances and should not balloon.
  return /[\p{L}\p{N}]{4,}/u.test(str);
}

function pdfFontInfo(
  fontName: string | undefined,
  styles: Record<string, { fontFamily?: string }>,
  fontSource?: PdfFontSourceLike,
): Pick<
  PdfTextItem,
  "fontFamily" | "displayFontFamily" | "sourceFontName" | "bold" | "italic"
> & { rawSourceFontName?: string } {
  const styleFamily = fontName ? styles[fontName]?.fontFamily : undefined;
  let sourceFontName = "";
  let rawSourceFontName: string | undefined;
  let bold = false;
  let italic = false;

  if (fontName) {
    try {
      const objs = fontSource?.commonObjs;
      if (objs?.has?.(fontName)) {
        const font = objs.get?.(fontName) as PdfFontObjectLike | undefined;
        const rawName = typeof font?.name === "string" ? font.name : undefined;
        // Keep the raw subset-tagged name (AAAABM+CambriaMath) for trust
        // lookups: each subset has its own charcode assignment, so the
        // normalized family name would conflate sibling subsets.
        rawSourceFontName = rawName || undefined;
        sourceFontName = normalizePdfSourceFontName(rawName);
        const lowerName = sourceFontName.toLowerCase();
        bold =
          Boolean(font?.black ?? font?.bold) ||
          /bold|black|heavy|semibold/.test(lowerName);
        italic = Boolean(font?.italic) || /italic|oblique/.test(lowerName);
      }
    } catch {
      // Some pdf.js font objects are lazily resolved; the style family is enough.
    }
  }

  const resolved = resolvePdfFontFamily({
    styleFamily,
    fontName: sourceFontName,
    bold,
    italic,
  });
  // pdf.js registers each embedded font as a FontFace under its loaded name
  // (g_dN_fM) while the document is open. Putting that name FIRST means the
  // live overlay renders the document's own font — the real fix for "the font
  // changed when I clicked" — with the classified stack as fallback. Only the
  // ephemeral display family carries it; the persisted stable family never
  // does (the FontFace dies with the document).
  const loadedFontFace =
    fontName && /^g_d\d+_f\d+$/i.test(fontName) ? fontName : null;
  const displayChain = pdfDisplayFontFamily(styleFamily, resolved.fontFamily);
  return {
    fontFamily: resolved.fontFamily,
    displayFontFamily: loadedFontFace
      ? `${loadedFontFace}, ${displayChain}`.slice(0, 400)
      : displayChain,
    sourceFontName: sourceFontName || undefined,
    rawSourceFontName,
    bold: resolved.bold,
    italic: resolved.italic,
  };
}

/**
 * Convert a pdf.js TextContent object into page-fraction text boxes. This is used
 * by both the read-only viewer (AI command context) and the editable text mode, so
 * BossAI sees the same ids/geometry/fonts that the user can later edit.
 */
export function textBoxesFromPdfTextContent(
  textContent: PdfTextContentLike,
  viewport: PdfViewportLike,
  idPrefix: string,
  fontSource?: PdfFontSourceLike,
): TextBox[] {
  const styles = textContent.styles ?? {};
  const items: PdfTextItem[] = textContentItems(textContent)
    .filter((raw): raw is PdfTextContentItemLike => {
      const item = raw as PdfTextContentItemLike | null;
      return typeof item?.str === "string" && isTextTransform(item.transform);
    })
    .map((item) => {
      const transform = (item.transform as number[]).slice(0, 6);
      const estimatedWidth = estimateItemWidth(
        item.str as string,
        transform,
        item.height,
      );
      const width =
        isFiniteNumber(item.width) && item.width > 0
          ? shouldUseEstimatedWidth(
              item.str as string,
              item.width,
              estimatedWidth,
            )
            ? estimatedWidth
            : item.width
          : estimatedWidth;
      const info = pdfFontInfo(
        typeof item.fontName === "string" ? item.fontName : undefined,
        styles,
        fontSource,
      );
      const mathFont = isMathSourceFontName(
        info.rawSourceFontName ?? info.sourceFontName,
      );
      const trustVerdict = mathFont
        ? textTrustForSourceFont(
            fontSource?.textTrust,
            info.rawSourceFontName,
            item.str as string,
          )
        : "unknown";
      return {
        str: item.str as string,
        transform,
        width,
        height: isFiniteNumber(item.height) ? item.height : undefined,
        fontFamily: info.fontFamily,
        displayFontFamily: info.displayFontFamily,
        sourceFontName: info.sourceFontName,
        bold: info.bold,
        italic: info.italic,
        kind: item.kind === "formula" ? "formula" : undefined,
        latex: typeof item.latex === "string" ? item.latex : undefined,
        mathFont: mathFont || undefined,
        untrustedText: trustVerdict === "untrusted" || undefined,
      };
    });
  return groupTextBoxes(items, viewport, idPrefix);
}
