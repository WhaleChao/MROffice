/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A slide show's faces, asked for once when the show starts -- not slide by
 * slide as the show reaches each one (TRIAGE 181).
 *
 * In the show the engine draws a slide's words once their faces' files are
 * read (the face metrics) and the browser draws them once the same slices are
 * loaded; a slide whose faces are not in yet is held (blank) up to 1.5 s and
 * then drawn the browser's way. The next slide is laid out ahead while this
 * one shows, which asks for its faces, but a flip faster than that lands on a
 * slide whose faces were never asked for -- a deck that changes face (a title
 * face, a CJK slice, a symbol face) on slide 30 held slide 30.
 *
 * Here the families the deck names (blocks', runs', table cells' and symbol
 * faces, `fontFamily` / `symbolFont` wherever the slide model holds them) and
 * the families the show's own first slide draws in (its stacks, the app's
 * fallbacks among them) are declared (`ensureOfficeFontFaces`), and every
 * declared slice that covers a character of the deck's text, at the weights
 * and styles the deck uses, is read for the engine (`faceMetrics`) and loaded
 * for the browser (FontFaceSet.load) -- all after the first slide is drawn
 * (its own faces it asks for itself): the next slide's first, the rest a few
 * at a time while the page is idle, so none competes with the first frame. The same entry points the slides call
 * themselves; nothing is decided here. Best effort: a failure is the slide's
 * own load's to meet, as before.
 *
 * The Word equivalent is lib/document-fonts-ahead.ts (the same declare, then
 * load, on a document's families).
 */
import { declaredFaces, faceCovers, selectComposite, splitFamilyList, type FaceDeclaration } from "@mrnote/core/web/lib/text-engine/face-index.ts";
import { faceMetricsReady } from "@mrnote/core/web/lib/text-engine/face-metrics.ts";

/** At most this many families are asked for (a deck's runs rarely name more). */
const MAX_FAMILIES = 32;
/** At most this many distinct characters (a CJK deck's text runs into thousands; the slices are per range). */
const MAX_CHARS = 8_000;
/** How long the declarations may take before the reads go out anyway. */
const DECLARE_WAIT_MS = 1_500;

const GENERIC = /^(serif|sans-serif|monospace|cursive|fantasy|system-ui|ui-\w+|inherit|initial|unset|emoji|math|-apple-system|blinkmacsystemfont)$/i;

export interface ShowFontNeeds {
  /** Family names, lower case (CSS matches them so), quotes off. */
  readonly families: string[];
  /** The deck's distinct characters. */
  readonly codePoints: number[];
  /** The (weight, italic) pairs the deck uses. */
  readonly styles: { weight: number; italic: boolean }[];
}

function addFamilies(value: unknown, into: Set<string>): void {
  if (typeof value !== "string") return;
  for (const family of splitFamilyList(value)) {
    const name = family.replace(/^["']|["']$/g, "").trim();
    if (name && name.length <= 64 && !GENERIC.test(name) && !name.startsWith("var(")) into.add(name);
  }
}

function weightOf(value: unknown): number | null {
  if (value === "bold" || value === "bolder") return 700;
  if (value === "normal") return 400;
  const number = typeof value === "number" ? value : typeof value === "string" ? Number(value) : NaN;
  return Number.isFinite(number) && number >= 1 && number <= 1000 ? number : null;
}

/** The families, characters and styles a deck's slides name (walked structurally: blocks, runs, cells, groups). */
export function showFontNeeds(slides: readonly unknown[], extraStacks: readonly string[] = []): ShowFontNeeds {
  const families = new Set<string>();
  const codePoints = new Set<number>();
  const styles = new Map<string, { weight: number; italic: boolean }>([["400/0", { weight: 400, italic: false }]]);
  const visit = (node: unknown, depth: number): void => {
    if (!node || typeof node !== "object" || depth > 48) return;
    if (Array.isArray(node)) {
      for (const item of node) visit(item, depth + 1);
      return;
    }
    const record = node as Record<string, unknown>;
    let weight: number | null = null;
    let italic = false;
    for (const [key, value] of Object.entries(record)) {
      if (key === "fontFamily" || key === "symbolFont" || key === "eastAsianFontFamily") addFamilies(value, families);
      else if (key === "text" && typeof value === "string" && codePoints.size < MAX_CHARS) {
        for (const char of value) {
          const code = char.codePointAt(0)!;
          if (code > 0x20) codePoints.add(code);
          if (codePoints.size >= MAX_CHARS) break;
        }
      } else if (key === "fontWeight" || key === "bold") weight = key === "bold" ? (value === true ? 700 : null) : weightOf(value);
      else if ((key === "fontStyle" && value === "italic") || (key === "italic" && value === true)) italic = true;
      else if (value && typeof value === "object") visit(value, depth + 1);
    }
    if (weight !== null || italic) styles.set(`${weight ?? 400}/${italic ? 1 : 0}`, { weight: weight ?? 400, italic });
  };
  visit(slides, 0);
  for (const stack of extraStacks) addFamilies(stack, families);
  return { families: [...families].slice(0, MAX_FAMILIES), codePoints: [...codePoints], styles: [...styles.values()] };
}







export function showFontSlices(needs: ShowFontNeeds): { face: FaceDeclaration; family: string; sample: string }[] {
  const out = new Map<string, { face: FaceDeclaration; family: string; sample: string }>();
  const declared = needs.families.map((family) => ({ family, faces: declaredFaces(family) })).filter((entry) => entry.faces.length > 0);
  for (const style of needs.styles) {
    const composites = declared.map(({ family, faces }) => ({ family, faces: selectComposite(faces, style.weight, style.italic) }));
    for (const code of needs.codePoints) {
      for (const { family, faces } of composites) {
        // The last declared face covering it wins within a family (CSS Fonts 4 §4.5: priority from the end).
        let face: FaceDeclaration | undefined;
        for (let i = faces.length - 1; i >= 0; i -= 1) if (faceCovers(faces[i]!, code)) { face = faces[i]; break; }
        if (!face) continue;
        if (face.url) {
          const known = out.get(face.url);
          if (!known) out.set(face.url, { face, family, sample: String.fromCodePoint(code) });
          else if (known.sample.length < 64 && !known.sample.includes(String.fromCodePoint(code))) known.sample += String.fromCodePoint(code);
        }
        break;
      }
    }
  }
  return [...out.values()];
}

/** Faces asked for at a time after the next slide's (the rest, while the page is idle). */
const BATCH = 4;
/** The longest the rest wait for the first slide to be drawn. */
const FIRST_SLIDE_WAIT_MS = 3_000;





const NOT_BEFORE_MS = 1_000;



function mark(name: string, at: number): void {
  if (typeof window === "undefined" || (window as { __mraiShowMarks?: boolean }).__mraiShowMarks !== true) return;
  try {
    performance.mark(name, { detail: { at } });
  } catch {
    // diagnostics only
  }
}

function load({ face, family, sample }: { face: FaceDeclaration; family: string; sample: string }): Promise<unknown> {
  const quoted = `"${family.replace(/["\\]/g, "")}"`;
  const weight = face.weightMin <= 400 && 400 <= face.weightMax ? 400 : face.weightMin;
  return Promise.allSettled([
    face.url ? faceMetricsReady([face.url]) : Promise.resolve(),
    document.fonts.load(`${face.style === "normal" ? "" : "italic "}${weight} 16px ${quoted}`, sample),
  ]);
}










export function warmShowFontsAhead(
  slides: readonly unknown[], stacks: () => readonly string[] = () => [], next: readonly unknown[] = [], firstDrawn: () => boolean = () => true,
): () => void {
  if (typeof document === "undefined" || !document.fonts) return () => undefined;
  let cancelled = false;
  const idle = () => new Promise((done) => (typeof window !== "undefined" && "requestIdleCallback" in window ? window.requestIdleCallback(done, { timeout: 1000 }) : setTimeout(done, 50)));
  void (async () => {
    const started = performance.now();
    // Drawn and still drawn a moment later (a slide's first commit can come before its bodies ask for faces), and
    // not before NOT_BEFORE_MS: the first slide's own font loads and its settling are the show's first frame.
    for (let steady = 0; !cancelled && (steady < 2 || performance.now() - started < NOT_BEFORE_MS) && performance.now() - started < FIRST_SLIDE_WAIT_MS; ) {
      await new Promise((done) => setTimeout(done, 100));
      steady = firstDrawn() ? steady + 1 : 0;
    }
    await idle();
    if (cancelled) return;
    const needs = showFontNeeds(slides, stacks());
    if (!needs.families.length || !needs.codePoints.length) return;
    const { ensureOfficeFontFaces, officeFontFacesSettled } = await import("@mrnote/core/web/lib/office-fonts/substitutes.ts");
    if (cancelled) return;
    for (const family of needs.families) ensureOfficeFontFaces(family);
    await officeFontFacesSettled(DECLARE_WAIT_MS);
    if (cancelled) return;
    const all = showFontSlices(needs);
    const nextCodes = new Set(showFontNeeds(next).codePoints);
    const isNext = (slice: { sample: string }) => [...slice.sample].some((char) => nextCodes.has(char.codePointAt(0)!));
    mark("mrai-show-faces-first", all.filter(isNext).length);
    for (const slice of all.filter(isNext)) void load(slice);
    const rest = all.filter((slice) => !isNext(slice));
    for (let i = 0; i < rest.length && !cancelled; i += BATCH) {
      await idle();
      if (cancelled) return;
      mark("mrai-show-faces-batch", i);
      await Promise.all(rest.slice(i, i + BATCH).map(load));
      mark("mrai-show-faces-batch-end", i);
    }
  })().catch(() => undefined);
  return () => {
    cancelled = true;
  };
}
