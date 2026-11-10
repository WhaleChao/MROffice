/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * How the PDF text editor names a font the PDF does not hold yet: a family
 * MRAI serves (lib/office-fonts/embeddable-faces.ts) asked for in a style.
 * The editor's draft keeps it where it keeps any font -- the draft's or a
 * part's `fontResourceName` -- so every bookkeeping rule a font change
 * follows (parts that split and join, typing that takes the format before it)
 * holds for it unchanged; the planner installs the face when it writes the
 * run (lib/pdf/native-embedded-fonts.ts) and names the installed resource.
 */

export const PDF_NATIVE_EMBEDDED_FONT_KEY_PREFIX = "mrai-embed:" as const;

export interface PdfNativeEmbeddedFontRequest {
  /** The family as the font list names it (an Office family). */
  readonly family: string;
  readonly bold: boolean;
  readonly italic: boolean;
}

export function pdfNativeEmbeddedFontKey(request: PdfNativeEmbeddedFontRequest): string {
  return `${PDF_NATIVE_EMBEDDED_FONT_KEY_PREFIX}${encodeURIComponent(request.family)}:${request.bold ? 1 : 0}:${request.italic ? 1 : 0}`;
}

export function isPdfNativeEmbeddedFontKey(value: unknown): value is string {
  return typeof value === "string" && value.startsWith(PDF_NATIVE_EMBEDDED_FONT_KEY_PREFIX);
}

/** The request a key names, or null when it is not one (or malformed). */
export function readPdfNativeEmbeddedFontKey(value: unknown): PdfNativeEmbeddedFontRequest | null {
  if (!isPdfNativeEmbeddedFontKey(value)) return null;
  const match = /^([^:]+):([01]):([01])$/u.exec(value.slice(PDF_NATIVE_EMBEDDED_FONT_KEY_PREFIX.length));
  if (!match) return null;
  let family: string;
  try {
    family = decodeURIComponent(match[1]);
  } catch {
    return null;
  }
  if (!family || family.length > 120) return null;
  return Object.freeze({ family, bold: match[2] === "1", italic: match[3] === "1" });
}

/** How the editor shows a run or part whose font is an embedded-font key:
 * as a font of the family it names, in its style, beside the PDF's own
 * fonts (the planner replaces it with the installed fonts it writes). */
export function pdfNativeEmbeddedFontDisplayOption<
  Option extends { readonly resourceName: string; readonly baseFont: string; readonly family: string; readonly bold: boolean; readonly italic: boolean },
>(key: string | null | undefined, like: Option | null | undefined): Option | undefined {
  const request = readPdfNativeEmbeddedFontKey(key);
  if (!request || !like) return undefined;
  return Object.freeze({
    ...like,
    resourceName: key as string,
    baseFont: request.family,
    family: request.family,
    bold: request.bold,
    italic: request.italic,
  });
}

/** Why a family could not be embedded, as the editor says it: the family
 * MRAI serves none of, the character its faces have no glyph for, a face
 * whose outlines cannot be embedded, a slice that could not be loaded, or
 * text a form draws (fonts are installed into a page's own text). */
export interface PdfNativeEmbeddedFontRefusal {
  readonly code: "unknown-family" | "missing-glyph" | "unsupported-face" | "unavailable" | "unsupported-place";
  readonly family: string;
  readonly character?: string;
}

const REFUSAL_TAG = /\[mrai-embedded-font ([a-z-]+) ([^\s\]]*) ([^\s\]]*)\]/u;
const REFUSAL_CODES = new Set<PdfNativeEmbeddedFontRefusal["code"]>([
  "unknown-family", "missing-glyph", "unsupported-face", "unavailable", "unsupported-place",
]);

/** The tag a refusal's message opens with, so it can be read back after
 * crossing the engine host as text. */
export function pdfNativeEmbeddedFontRefusalTag(refusal: PdfNativeEmbeddedFontRefusal): string {
  return `[mrai-embedded-font ${refusal.code} ${encodeURIComponent(refusal.family)} ${encodeURIComponent(refusal.character ?? "")}]`;
}

/** The refusal an error carries (the planner's own, or one rebuilt from an
 * engine answer), or null. */
export function readPdfNativeEmbeddedFontRefusal(error: unknown): PdfNativeEmbeddedFontRefusal | null {
  const message = error instanceof Error ? error.message : typeof error === "string" ? error : "";
  const match = REFUSAL_TAG.exec(message);
  if (!match || !REFUSAL_CODES.has(match[1] as PdfNativeEmbeddedFontRefusal["code"])) return null;
  try {
    const family = decodeURIComponent(match[2]);
    const character = decodeURIComponent(match[3]);
    return Object.freeze({
      code: match[1] as PdfNativeEmbeddedFontRefusal["code"],
      family,
      ...(character ? { character } : {}),
    });
  } catch {
    return null;
  }
}

/** Whether a draft (or a part of it) sets text in a family the PDF does not
 * hold: its writes install fonts, through the font-extension route. */
export function pdfNativeTextDraftEmbedsFonts(
  draft: Readonly<{ fontResourceName: string; segments?: readonly Readonly<{ fontResourceName?: string }>[] }> | null | undefined,
): boolean {
  return Boolean(draft && (
    isPdfNativeEmbeddedFontKey(draft.fontResourceName) ||
    (draft.segments ?? []).some((segment) => isPdfNativeEmbeddedFontKey(segment.fontResourceName))
  ));
}

/** Every family-and-style a draft sets text in that the PDF does not hold. */
export function pdfNativeTextDraftEmbeddedFontRequests(
  draft: Readonly<{ fontResourceName: string; segments?: readonly Readonly<{ fontResourceName?: string }>[] }> | null | undefined,
): readonly PdfNativeEmbeddedFontRequest[] {
  const keys = new Set<string>();
  if (draft && isPdfNativeEmbeddedFontKey(draft.fontResourceName)) keys.add(draft.fontResourceName);
  for (const segment of draft?.segments ?? []) {
    if (isPdfNativeEmbeddedFontKey(segment.fontResourceName)) keys.add(segment.fontResourceName);
  }
  return [...keys].flatMap((key) => {
    const request = readPdfNativeEmbeddedFontKey(key);
    return request ? [request] : [];
  });
}
