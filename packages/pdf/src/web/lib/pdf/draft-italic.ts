/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { PDF_NATIVE_SYNTHETIC_ITALIC_SKEW } from "./native-synthetic-style";

/** The writer's slant, as the angle CSS `skewX()` and the raster shear take. */
export const PDF_DRAFT_SYNTHETIC_ITALIC_SKEW_DEG =
  (Math.atan(PDF_NATIVE_SYNTHETIC_ITALIC_SKEW) * 180) / Math.PI;

/** The CSS transform that slants upright glyphs as the writer does. */
export function pdfDraftSyntheticSlantTransform(): string {
  return `skewX(${(-PDF_DRAFT_SYNTHETIC_ITALIC_SKEW_DEG).toFixed(3)}deg)`;
}

export interface PdfDraftItalicPresentation {
  /** What `font-style` the draft, the input and the measuring canvas ask for. */
  readonly fontStyle: "italic" | "normal";
  /** The glyphs are upright and the slant is drawn as a shear. */
  readonly synthetic: boolean;
}

interface FontFaceLike {
  readonly family: string;
  readonly style: string;
}

interface DocumentWithFonts {
  readonly fonts?: Iterable<FontFaceLike>;
}

function firstFamily(fontFamily: string): string {
  const first = fontFamily.split(",")[0] ?? "";
  return first.trim().replace(/^["']|["']$/g, "").trim();
}

/**
 * Whether the first family of `fontFamily` has an italic face among the
 * faces the document registered: "real" when one is italic or oblique,
 * "none" when the family is registered upright only, "unknown" when the
 * document registered no face of that name (the reader's own font).
 */
export function cssFamilyItalicFace(
  fontFamily: string,
  doc: DocumentWithFonts | undefined = typeof document === "undefined" ? undefined : document,
): "real" | "none" | "unknown" {
  const family = firstFamily(fontFamily);
  if (!family || !doc?.fonts) return "unknown";
  let registered = false;
  try {
    for (const face of doc.fonts) {
      if (face.family.replace(/^["']|["']$/g, "") !== family) continue;
      registered = true;
      if (/^(italic|oblique)/i.test(String(face.style))) return "real";
    }
  } catch {
    return "unknown";
  }
  return registered ? "none" : "unknown";
}

/** The presentation of `italic` over `fontFamily` (see the module note). */
export function pdfDraftItalic(
  fontFamily: string,
  italic: boolean,
  doc?: DocumentWithFonts,
): PdfDraftItalicPresentation {
  if (!italic) return { fontStyle: "normal", synthetic: false };
  const synthetic = cssFamilyItalicFace(fontFamily, doc) === "none";
  return { fontStyle: synthetic ? "normal" : "italic", synthetic };
}












export function slantPdfDraftWords(element: HTMLElement, baselinePx: number): void {
  const doc = element.ownerDocument;
  const view = doc.defaultView;
  const texts: Text[] = [];
  const walker = doc.createTreeWalker(element, 4 /* NodeFilter.SHOW_TEXT */);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) texts.push(node as Text);
  const transform = pdfDraftSyntheticSlantTransform();
  const origin = `0 ${Math.round(baselinePx * 100) / 100}px`;
  const rulesOf = (start: Node | null): string => {
    const lines = new Set<string>();
    for (let at = start; at && at !== element.parentNode; at = at.parentNode) {
      if (!(at instanceof (view?.HTMLElement ?? HTMLElement))) continue;
      const own = at.style.textDecorationLine || (view ? view.getComputedStyle(at).textDecorationLine : "");
      for (const line of own.split(/\s+/u)) if (line && line !== "none") lines.add(line);
      if (at === element) break;
    }
    return [...lines].join(" ");
  };
  for (const text of texts) {
    const value = text.data;
    if (!/\S/u.test(value)) continue;
    const rules = rulesOf(text.parentNode);
    const parts: Node[] = [];
    for (const token of value.match(/\S+|\s+/gu) ?? []) {
      if (!/\S/u.test(token)) {
        parts.push(doc.createTextNode(token));
        continue;
      }
      const word = doc.createElement("span");
      word.dataset.pdfDraftSlantedWord = "true";
      word.style.display = "inline-block";
      word.style.transformOrigin = origin;
      word.style.transform = transform;
      if (rules) word.style.textDecorationLine = rules;
      word.textContent = token;
      parts.push(word);
    }
    text.replaceWith(...parts);
  }
}
