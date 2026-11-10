/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What every editor adapter shares on the page: the switch, a canvas `font`
 * string resolved to the engine's font (CSS custom properties in the family
 * list are resolved against the root, as the browser does for elements), and
 * the bridge that tells font-keyed caches when the engine's faces land (a
 * `loadingdone` on document.fonts, the event they already listen for).
 */
import { textEngineEnabled, type TextEngineSurface } from "./flags";
import type { KinsokuList } from "./kinsoku";
import { breakLines, lineBreakOpportunities } from "./line-break";
import { engineFontFromCss, explainText, measureText, onTextEngineChange, type EngineFont } from "./measure";

let bridged = false;
function bridgeFontEvents(): void {
  if (bridged || typeof document === "undefined" || !document.fonts) return;
  bridged = true;
  // Probes compare the editors' engine answers across browsers through this.
  (window as unknown as Record<string, unknown>).__mraiPageTextEngine = { pageTextWidth, pageWrappedLineCount, measureText, explainText, pageEngineFont };
  let queued = false;
  onTextEngineChange(() => {
    if (queued) return;
    queued = true;
    setTimeout(() => {
      queued = false;
      try {
        document.fonts.dispatchEvent(new Event("loadingdone"));
      } catch { /* no EventTarget fonts */ }
    }, 0);
  });
}

/** Editors whose text the browser draws with its default kerning (the PDF
 * overlays set no font-kerning): measured kerned, as drawn. */
const KERNED_SURFACES: ReadonlySet<TextEngineSurface> = new Set(["pdf"]);

export function textEngineOn(surface: TextEngineSurface): boolean {
  if (!textEngineEnabled(surface)) return false;
  bridgeFontEvents();
  return true;
}

/** `var(--name, fallback)` in a family list, resolved against the root. */
export function resolveFamilyVariables(stack: string): string {
  if (!stack.includes("var(") || typeof document === "undefined") return stack;
  const style = getComputedStyle(document.documentElement);
  let out = stack;
  for (let guard = 0; guard < 8 && out.includes("var("); guard += 1) {
    out = out.replace(/var\(\s*(--[\w-]+)\s*(?:,\s*([^()]*(?:\([^()]*\)[^()]*)*))?\)/g, (_, name: string, fallback?: string) => {
      const value = style.getPropertyValue(name).trim();
      return value || (fallback ?? "").trim();
    });
  }
  return out;
}

/** The engine font of a canvas `font` string, variables resolved. */
export function pageEngineFont(fontCss: string): EngineFont | undefined {
  const font = engineFontFromCss(fontCss);
  return font ? { ...font, stack: resolveFamilyVariables(font.stack) } : undefined;
}

/** A text's width in a canvas `font` string, from the served files, or
 * undefined (engine off for the surface, faces loading, not measurable). */
export function pageTextWidth(surface: TextEngineSurface, text: string, fontCss: string, kerning = KERNED_SURFACES.has(surface)): number | undefined {
  if (!textEngineOn(surface)) return undefined;
  const font = pageEngineFont(fontCss);
  if (!font) return undefined;
  const measured = measureText(text, { ...font, kerning });
  return measured.status === "ok" ? measured.width : undefined;
}

/**
 * The lines of a text wrapped at `widthPx` the way CSS `pre-wrap` +
 * `overflow-wrap: break-word` wraps it (hard breaks at "\n"), from the served
 * faces: the line count every browser agrees on. Undefined when the engine
 * cannot measure every character.
 */
export function pageWrappedLineCount(
  surface: TextEngineSurface, text: string, fontCss: string, widthPx: number, kinsoku: KinsokuList | null = null,
): number | undefined {
  if (!textEngineOn(surface)) return undefined;
  const plain = pageEngineFont(fontCss);
  if (!plain) return undefined;
  const font = { ...plain, kerning: KERNED_SURFACES.has(surface) };
  let lines = 0;
  for (const paragraph of text.split(/\r\n|\r|\n/)) {
    if (!paragraph) {
      lines += 1;
      continue;
    }
    const measured = measureText(paragraph, font);
    if (measured.status !== "ok") return undefined;
    const graphemes = measured.graphemes;
    const opportunities = lineBreakOpportunities(graphemes.map((grapheme) => grapheme.text), { kinsoku });
    lines += breakLines(graphemes, widthPx, opportunities).length;
  }
  return lines;
}

/**
 * The height a laid-out text element would take if its lines came from the
 * engine instead of the browser: its DOM height with each "text holder"
 * (an element whose children are text runs) re-measured -- the holder's own
 * width, the engine's line breaks over its runs in their computed fonts,
 * times its line height. Everything else (insets, paragraph spacing,
 * markers) keeps its CSS values, which are the same in every browser.
 * Undefined when any holder holds something the engine does not lay out
 * (pictures, tabs, runs of differing line heights, `normal` line height).
 */
export function pageEngineBlockHeight(surface: TextEngineSurface, frame: HTMLElement, holderSelector: string): number | undefined {
  if (!textEngineOn(surface)) return undefined;
  const view = frame.ownerDocument.defaultView;
  if (!view) return undefined;
  let height = frame.getBoundingClientRect().height;
  for (const holder of Array.from(frame.querySelectorAll<HTMLElement>(holderSelector))) {
    const holderStyle = view.getComputedStyle(holder);
    const lineHeight = Number.parseFloat(holderStyle.lineHeight);
    if (!Number.isFinite(lineHeight) || !holderStyle.lineHeight.endsWith("px")) return undefined;
    const width = holder.getBoundingClientRect().width;
    // Runs, in order: [text, font, kerning]; hard breaks split lines.
    const paragraphs: { text: string; font: EngineFont }[][] = [[]];
    let unsupported = false;
    const walk = (node: Node): void => {
      if (unsupported) return;
      if (node.nodeType === Node.TEXT_NODE) {
        const parent = node.parentElement ?? holder;
        const style = view.getComputedStyle(parent);
        if (style.lineHeight !== holderStyle.lineHeight && parent !== holder) {
          const own = Number.parseFloat(style.lineHeight);
          if (!(own <= lineHeight)) { unsupported = true; return; }
        }
        if (style.letterSpacing !== "normal" && Number.parseFloat(style.letterSpacing) !== 0) { unsupported = true; return; }
        const font = pageEngineFont(`${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`);
        if (!font) { unsupported = true; return; }
        const kerned: EngineFont = { ...font, kerning: style.fontKerning !== "none" };
        const pieces = (node.textContent ?? "").split(/\r\n|\r|\n/);
        pieces.forEach((piece, index) => {
          if (index > 0) paragraphs.push([]);
          if (piece) paragraphs[paragraphs.length - 1].push({ text: piece, font: kerned });
        });
        return;
      }
      if (!(node instanceof HTMLElement)) return;
      if (node.tagName === "BR") { paragraphs.push([]); return; }
      if (!["SPAN", "A", "B", "I", "U", "S", "STRONG", "EM", "SUB", "SUP"].includes(node.tagName)) { unsupported = true; return; }
      if (["SUB", "SUP"].includes(node.tagName) || view.getComputedStyle(node).display !== "inline") { unsupported = true; return; }
      node.childNodes.forEach(walk);
    };
    holder.childNodes.forEach(walk);
    if (unsupported) return undefined;
    let lines = 0;
    for (const runs of paragraphs) {
      if (!runs.length) { lines += 1; continue; }
      const graphemes: { text: string; advance: number }[] = [];
      for (const run of runs) {
        const measured = measureText(run.text, run.font);
        if (measured.status !== "ok") return undefined;
        graphemes.push(...measured.graphemes);
      }
      const opportunities = lineBreakOpportunities(graphemes.map((grapheme) => grapheme.text), { kinsoku: null });
      lines += holderStyle.whiteSpace === "pre" ? 1 : breakLines(graphemes, width, opportunities).length;
    }
    height += lines * lineHeight - holder.getBoundingClientRect().height;
  }
  return height;
}
