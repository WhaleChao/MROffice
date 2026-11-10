/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export function pdfJsFaceDrawsText(input: Readonly<{
  document: Document | null | undefined;
  family: string;
  text: string;
  /** The size the page sets the line at, and the width it sets it in, px. */
  fontSizePx: number;
  widthPx: number;
}>): boolean {
  const text = input.text.trim();
  if (!input.document || !text || !(input.fontSizePx > 0) || !/^g_d\d+_f\d+(?:-mrai\d*)?$/u.test(input.family)) return true;
  let context: CanvasRenderingContext2D | null = null;
  try { context = input.document.createElement("canvas").getContext("2d"); } catch { context = null; }
  if (!context || typeof context.measureText !== "function") return true;
  const measure = (fallback: string) => {
    context!.font = `${input.fontSizePx}px "${input.family}", ${fallback}`;
    return context!.measureText(text).width;
  };
  const mono = measure("monospace"), serif = measure("serif");
  if (!(mono > 0) || !(serif > 0)) return true;
  // A character the face does not hold is drawn by the fallback (the two
  // fallbacks give two widths): the face draws what it holds, the stack the
  // rest, and its width says nothing of it.
  if (Math.abs(mono - serif) > Math.max(0.5, mono * 0.002)) return true;
  // Held, but other glyphs than the page's: another width.
  return !(input.widthPx > 0) || Math.abs(mono - input.widthPx) / input.widthPx <= 0.12;
}

/** A font stack without the faces pdf.js loaded that `keep` refuses. */
export function withoutPdfJsFaces(stack: string, keep: (family: string) => boolean): string {
  const kept = stack.split(",").filter((token) => {
    const family = token.trim().replace(/^["']|["']$/gu, "");
    return !/^g_d\d+_f\d+$/u.test(family) || keep(family);
  });
  return kept.length ? kept.join(",").trimStart() : stack;
}

/**
 * The weight a line drafted in other faces than its own is set at, read by
 * width: the line's own text measured in the stack regular and bold, the one
 * nearer the width the page sets it in. A scrambled face names nothing of its
 * weight (the exam cover's bold "AllAndNone" drafted regular). Null where it
 * cannot be measured, or neither is near.
 */
export function pdfDraftWeightByWidth(input: Readonly<{
  document: Document | null | undefined;
  stack: string;
  text: string;
  fontSizePx: number;
  widthPx: number;
}>): "400" | "700" | null {
  const text = input.text.trim();
  if (!input.document || !text || !(input.fontSizePx > 0) || !(input.widthPx > 0)) return null;
  let context: CanvasRenderingContext2D | null = null;
  try { context = input.document.createElement("canvas").getContext("2d"); } catch { context = null; }
  if (!context || typeof context.measureText !== "function") return null;
  const measure = (weight: string) => {
    context!.font = `${weight} ${input.fontSizePx}px ${input.stack}`;
    return context!.measureText(text).width;
  };
  const regular = measure("400"), bold = measure("700");
  if (!(regular > 0) || !(bold > 0) || Math.abs(bold - regular) < regular * 0.02) return null;
  const nearer = Math.abs(bold - input.widthPx) < Math.abs(regular - input.widthPx) ? "700" : "400";
  const off = Math.abs((nearer === "700" ? bold : regular) - input.widthPx) / input.widthPx;
  return off <= 0.12 ? nearer : null;
}

/**
 * A face's weight and slant as its PostScript name writes them, run together
 * as producers do ("FrutigerLTPro-BoldCnIt", "MinionPro-It",
 * "Helvetica-BoldOblique"). The copy of a face registered for drafts and the
 * draft's own weight and slant read them alike: a face registered regular and
 * asked bold (or upright and asked italic) is bolded (slanted) a second time
 * by the browser.
 */
export function pdfFaceNameStyle(name: string | null | undefined): Readonly<{ bold: boolean; italic: boolean }> {
  const face = (name ?? "").replace(/^[A-Z]{6}\+/u, "");
  return {
    bold: /bold|black|heavy|demi|(?:^|[-_ ,])(?:bd|blk|hv)(?:cn|it)?$|[a-z](?:bd|blk)(?:cn)?(?:it)?$/iu.test(face),
    italic: /italic|oblique|slanted|(?:^|[-_ ,])(?:it|ita|obl)$|(?:cn|cond|bd|bold|lt|light|md|medium|bk|book|rg|regular|blk|black|hv|heavy|semibold|roman)it$/iu.test(face),
  };
}
