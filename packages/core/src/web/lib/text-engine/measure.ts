/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Text measurement from the served font files: the advance of every
 * grapheme of a text in a CSS font, as the face the browser draws it in
 * stores it -- no kerning, no ligatures, no hinting, no platform rounding.
 * Office sets cell and body text that way unless a document turns kerning
 * on; callers apply Office's own device rounding on top.
 *
 * Answers synchronously. A text whose faces are still loading answers
 * "pending" (the caller keeps its browser measurement and hears
 * `onTextEngineChange` when the faces land); a text the engine cannot
 * measure faithfully (a system font, a script that needs shaping, an
 * installed local font that differs from the served file) answers
 * "unsupported".
 */
import { declaredFaces, faceCovers, faceIndexSignature, GENERIC_FAMILIES, selectComposite, splitFamilyList, type FaceDeclaration } from "./face-index";
import { faceMetrics, faceMetricsGeneration, onFaceMetricsLoaded, type FaceMetrics } from "./face-metrics";
import { syntheticBoldExtraPx } from "./synthetic-bold";
import { glyphAdvanceUnits } from "./variation";
import { pairKerning } from "./kerning";
import { officePairKerning } from "./office-kerning";
import { apiFaceOrder, onApiFaceOrderKnown } from "./face-order";
import { faceVerification, faceVerificationDetail, faceVerificationGeneration, installedTwinVerification, localOnlyFaceState, onFaceVerified, systemFamilyState } from "./verify";

export interface EngineFont {
  /** CSS font-family list. */
  readonly stack: string;
  readonly weight: number;
  readonly italic: boolean;
  readonly sizePx: number;
  /** Apply the faces' pair kerning (PowerPoint `kern`, Word `w:kern`);
   * Excel and Word body text by default do not. */
  readonly kerning?: boolean;
  /** Follow a local()-only face that is not installed to the family's
   * earlier faces, as the browser does (Docs); otherwise such a face leaves
   * the text to the browser. */
  readonly followLocalOnly?: boolean;
  /** Read a face the browser draws from an installed font that differs
   * from the served file with the served file's metrics, as an estimate
   * (the measurement says `approximate`), instead of leaving it to the
   * browser: a plan over pages not laid out yet, which the page corrects. */
  readonly approximateInstalled?: boolean;
}

export interface MeasuredGrapheme {
  readonly text: string;
  /** Advance in CSS px at the font's size. */
  readonly advance: number;
  /** Drawn in a bold the face does not have (synthesized). */
  readonly syntheticBold?: true;
}

export type Measurement =
  | { readonly status: "ok"; readonly width: number; readonly graphemes: readonly MeasuredGrapheme[]; readonly approximate?: true }
  | { readonly status: "pending" }
  | { readonly status: "unsupported"; readonly reason: string };

const PENDING = { status: "pending" } as const;

/** Scripts that need shaping (joining, reordering, marks positioned by GPOS)
 * and the bidi controls: their width is not the sum of advances. */
const SHAPED = /[\u0590-\u08ff\u0900-\u0dff\u0e00-\u0fffက-႟ក-\u18afꠀ-\uabffיִ-﷿ﹰ-\ufefe\u200d\u200e\u200f\u202a-\u202e\u2066-\u2069\ufe0f\t]|[\u{10000}-\u{1ffff}]/u;
/** Code points that draw nothing and take no room. */
const IGNORABLE = new Set([0x200b, 0x200c, 0x2060, 0xfeff, 0xad, 0xfe00, 0xfe01, 0xfe02, 0xfe03, 0xfe0e]);

interface ResolvedGrapheme {
  /** Advance as a fraction of the em, size-adjust applied. */
  readonly em: number;
  readonly syntheticBold: boolean;
  /** The face's metrics and the first/last glyph, for pair kerning. */
  readonly metrics?: FaceMetrics;
  readonly scale?: number;
  readonly firstGlyph?: number;
  readonly lastGlyph?: number;
  /** The face and the first/last code point, for Office's kerning of a stand-in (office-kerning.ts). */
  readonly face?: FaceDeclaration;
  readonly firstCode?: number;
  readonly lastCode?: number;
  /** Read from a served file the installed font differs from. */
  readonly approximate?: true;
}
type GraphemeAnswer = ResolvedGrapheme | "pending" | string;

let epoch = "";
const graphemeCache = new Map<string, Map<string, GraphemeAnswer>>();
const textCache = new Map<string, Measurement>();
const TEXT_CACHE_LIMIT = 50_000;

function currentEpoch(): string {
  return `${faceIndexSignature()}|${faceMetricsGeneration()}|${faceVerificationGeneration()}|${apiFaceOrder()}`;
}

/** The face a font draws a code point in, as `measureLineBoxHeight` resolved it, until the epoch moves. */
const lineFaces = new Map<string, Map<number, Exclude<ReturnType<typeof resolveFace>, string>>>();

function refreshEpoch(): void {
  const now = currentEpoch();
  if (now === epoch) return;
  epoch = now;
  graphemeCache.clear();
  textCache.clear();
  lineFaces.clear();
}

/** A token that changes whenever an answer may have changed: fold it into
 * any cache built on measurements. */
export function textEngineEpoch(): string {
  refreshEpoch();
  return epoch;
}

/**
 * A token that changes only when an answer given may have changed: the
 * faces declared and the order the browser takes them in. A face file
 * loading or a face verified changes only answers that were pending (every
 * face a given answer read had loaded and was verified), so a cache of
 * answers that were given may key on this instead of the epoch.
 */
export function textEngineAnswerEpoch(): string {
  return `${faceIndexSignature()}|${apiFaceOrder()}`;
}

/** Hear when faces load or are verified (an answer may have moved from
 * pending). */
export function onTextEngineChange(listener: () => void): () => void {
  const a = onFaceMetricsLoaded(listener);
  const b = onFaceVerified(listener);
  const c = onApiFaceOrderKnown(listener);
  return () => { a(); b(); c(); };
}

/** The face that draws `codePoint` in a family list, per CSS font fallback. */
export function resolveFace(
  families: readonly string[], weight: number, italic: boolean, codePoint: number, followLocalOnly = false, approximateInstalled = false,
): { face: FaceDeclaration; metrics: FaceMetrics; glyph: number; syntheticBold: boolean; approximate?: true } | "pending" | string {
  let installedSkipped = false;
  for (const family of families) {
    if (GENERIC_FAMILIES.has(family)) return `generic family ${family}`;
    const faces = declaredFaces(family);
    // An undeclared name is a system font, or nothing: the engine cannot
    // know what an installed one draws. Docs (`followLocalOnly`) asks the
    // browser whether the name draws anything; one that does not is passed
    // over, as the browser passes over it.
    if (!faces.length) {
      if (!followLocalOnly) return `undeclared family ${family}`;
      const system = systemFamilyState(family);
      if (system === "absent") continue;
      // An installed system font the engine has no file of: an estimate
      // reads the character from the families after it.
      if (approximateInstalled) { installedSkipped = true; continue; }
      return `undeclared family ${family}`;
    }
    const order = apiFaceOrder();
    if (order === "pending") return "pending";
    const composite = selectComposite(faces, weight, italic, order);
    // An installed font a local()-only face draws, with no file of its own
    // to read: the family's served faces stand for it when they are its
    // metric twin (an installed Calibri and the served Carlito).
    let installed = false;
    for (let index = composite.length - 1; index >= 0; index -= 1) {
      const face = composite[index];
      if (!faceCovers(face, codePoint)) continue;
      if (!face.url) {
        // A local()-only face: drawn only when the font is installed;
        // otherwise its characters fall through to the earlier faces.
        if (!face.localNames.length || !followLocalOnly) return `face of ${family} has no served file`;
        const local = localOnlyFaceState(face);
        if (local === "pending") return "pending";
        if (local === "missing") continue;
        // No file to read the installed font from: an estimate reads the
        // character from the faces after it, as if it were not installed.
        if (approximateInstalled) installedSkipped = true;
        else installed = true;
        continue;
      }
      const metrics = faceMetrics(face.url);
      if (metrics === "pending") return "pending";
      if (metrics === "failed") return `face file failed: ${face.url}`;
      const glyph = metrics.cmap.get(codePoint);
      // In range but not in the file: the browser tries the next face, then
      // the next family.
      if (glyph === undefined) continue;
      if (installed) {
        // The installed font draws the character (or, lacking it, this
        // face does): the file's widths are the drawn ones only when the
        // installed font measures as the file does.
        const twin = installedTwinVerification(face, metrics);
        if (twin === "pending") return "pending";
        if (twin === "differs") return `installed font: ${family}`;
        return { face, metrics, glyph, syntheticBold: weight >= 600 && face.weightMax < 600 };
      }
      const verified = faceVerification(face);
      if (verified === "pending") return "pending";
      if (verified === "differs" && !approximateInstalled) return `installed font differs: ${family}`;
      return { face, metrics, glyph, syntheticBold: weight >= 600 && face.weightMax < 600, ...(verified === "differs" || installedSkipped ? { approximate: true as const } : {}) };
    }
    // The installed font may have the character no served face has.
    if (installed) return `installed font: ${family}`;
  }
  return "no face covers the character";
}

function resolveGrapheme(families: readonly string[], weight: number, italic: boolean, grapheme: string, followLocalOnly = false, approximateInstalled = false): GraphemeAnswer {
  if (SHAPED.test(grapheme)) return "needs shaping";
  let em = 0;
  let syntheticBold = false;
  let approximate = false;
  let base: { face: FaceDeclaration; metrics: FaceMetrics; glyph: number } | null = null;
  let lastGlyph: number | undefined;
  let firstCode: number | undefined;
  let lastCode: number | undefined;
  for (const char of grapheme) {
    let codePoint = char.codePointAt(0)!;
    if (IGNORABLE.has(codePoint)) continue;
    if (codePoint < 0x20 || (codePoint >= 0x7f && codePoint < 0xa0)) return "control character";
    // Both engines draw a no-break space with the space glyph.
    if (codePoint === 0xa0) codePoint = 0x20;
    if (base) {
      // A combining mark is drawn from its base's face when it has it.
      const glyph = base.metrics.cmap.get(codePoint);
      if (glyph === undefined) return "mark outside its base's face";
      const units = glyphAdvanceUnits(base.face, base.metrics, glyph, weight);
      if (!Number.isFinite(units)) return "variable face the engine does not place";
      em += units / base.metrics.unitsPerEm * base.face.sizeAdjust;
      lastGlyph = glyph;
      lastCode = codePoint;
      continue;
    }
    const resolved = resolveFace(families, weight, italic, codePoint, followLocalOnly, approximateInstalled);
    if (typeof resolved === "string") return resolved;
    base = resolved;
    if (resolved.approximate) approximate = true;
    syntheticBold = resolved.syntheticBold;
    const units = glyphAdvanceUnits(resolved.face, resolved.metrics, resolved.glyph, weight);
    if (!Number.isFinite(units)) return "variable face the engine does not place";
    em += units / resolved.metrics.unitsPerEm * resolved.face.sizeAdjust;
    lastGlyph = resolved.glyph;
    firstCode = lastCode = codePoint;
  }
  return base
    ? { em, syntheticBold, metrics: base.metrics, scale: base.face.sizeAdjust / base.metrics.unitsPerEm, firstGlyph: base.glyph, lastGlyph, face: base.face, firstCode, lastCode, ...(approximate ? { approximate: true as const } : {}) }
    : { em, syntheticBold };
}

let segmenter: Intl.Segmenter | undefined;
export function graphemesOf(text: string): string[] {
  // ASCII is one grapheme per character; skip the segmenter for it.
  if (/^[\x20-\x7e]*$/.test(text)) return text.split("");
  segmenter ??= new Intl.Segmenter(undefined, { granularity: "grapheme" });
  return Array.from(segmenter.segment(text), (part) => part.segment);
}

/** Advances of every grapheme of `text` in `font`. */
export function measureText(text: string, font: EngineFont): Measurement {
  refreshEpoch();
  const fontKey = `${font.stack}\u0000${font.weight}\u0000${font.italic ? 1 : 0}${font.followLocalOnly ? "\u0000l" : ""}${font.approximateInstalled ? "\u0000a" : ""}`;
  const kerningKey = font.kerning ? "\u0000k" : "";
  const textKey = `${fontKey}\u0000${font.sizePx}${kerningKey}\u0000${text}`;
  const known = textCache.get(textKey);
  if (known) return known;
  let perFont = graphemeCache.get(fontKey);
  if (!perFont) graphemeCache.set(fontKey, (perFont = new Map()));
  const families = splitFamilyList(font.stack);
  const graphemes: MeasuredGrapheme[] = [];
  let width = 0;
  let pending = false;
  let approximate = false;
  const boldExtra = syntheticBoldExtraPx(font.sizePx);
  let previous: ResolvedGrapheme | undefined;
  for (const grapheme of graphemesOf(text)) {
    let answer = perFont.get(grapheme);
    if (answer === undefined) {
      answer = resolveGrapheme(families, font.weight, font.italic, grapheme, font.followLocalOnly, font.approximateInstalled);
      perFont.set(grapheme, answer);
    }
    if (answer === "pending") {
      pending = true;
      continue;
    }
    if (typeof answer === "string") {
      const result: Measurement = { status: "unsupported", reason: answer };
      remember(textKey, result);
      return result;
    }
    const advance = answer.em * font.sizePx + (answer.syntheticBold ? boldExtra : 0);
    // Pair kerning moves the previous glyph's advance, within one face.
    if (font.kerning && previous?.metrics && previous.metrics === answer.metrics &&
        previous.lastGlyph !== undefined && answer.firstGlyph !== undefined && graphemes.length) {
      // A served stand-in of an Office family kerns by Office's pairs.
      const office = previous.face && previous.lastCode !== undefined && answer.firstCode !== undefined
        ? officePairKerning(previous.face, font.weight, font.italic, previous.lastCode, answer.firstCode)
        : null;
      if (office === "pending") pending = true;
      const em = office === "pending" ? 0 : office ?? (pairKerning(previous.metrics)?.(previous.lastGlyph, answer.firstGlyph) ?? 0) * (previous.scale ?? 0);
      if (em) {
        const delta = em * font.sizePx;
        const last = graphemes[graphemes.length - 1];
        graphemes[graphemes.length - 1] = { ...last, advance: last.advance + delta };
        width += delta;
      }
    }
    previous = answer;
    if (answer.approximate) approximate = true;
    graphemes.push(answer.syntheticBold ? { text: grapheme, advance, syntheticBold: true } : { text: grapheme, advance });
    width += advance;
  }
  if (pending) return PENDING;
  const result: Measurement = approximate ? { status: "ok", width, graphemes, approximate: true } : { status: "ok", width, graphemes };
  remember(textKey, result);
  return result;
}

function remember(key: string, value: Measurement): void {
  if (textCache.size >= TEXT_CACHE_LIMIT) textCache.clear();
  textCache.set(key, value);
}

/** The width of `text`, or undefined when the engine cannot answer yet. */
export function measureTextWidth(text: string, font: EngineFont): number | undefined {
  const measured = measureText(text, font);
  return measured.status === "ok" ? measured.width : undefined;
}

export interface EngineLineMetrics {
  /** px, positive up. */
  readonly ascent: number;
  /** px, positive down. */
  readonly descent: number;
  readonly lineGap: number;
}

/**
 * The line metrics of a font's first available face (the face that draws
 * U+0020, as CSS defines the primary font), on macOS terms: the declared
 * overrides, else the hhea values.
 */
export function measureLineMetrics(font: EngineFont): EngineLineMetrics | "pending" | "unsupported" {
  refreshEpoch();
  const resolved = resolveFace(splitFamilyList(font.stack), font.weight, font.italic, 0x20, font.followLocalOnly, font.approximateInstalled);
  if (resolved === "pending") return "pending";
  if (typeof resolved === "string") return "unsupported";
  const { face, metrics } = resolved;
  const unit = font.sizePx * face.sizeAdjust;
  return {
    ascent: (face.ascentOverride ?? metrics.hhea.ascender / metrics.unitsPerEm) * unit,
    descent: (face.descentOverride ?? -metrics.hhea.descender / metrics.unitsPerEm) * unit,
    lineGap: (face.lineGapOverride ?? metrics.hhea.lineGap / metrics.unitsPerEm) * unit,
  };
}

/**
 * A canvas/CSS `font` shorthand as the editors build it:
 * `[style] [weight] <size>px|pt <family list>`. Undefined for anything else.
 */
export function engineFontFromCss(css: string): EngineFont | undefined {
  const match = /^\s*(?:(normal|italic|oblique)\s+)?(?:(normal|bold|bolder|lighter|[1-9]00|[0-9]+(?:\.[0-9]+)?)\s+)?([0-9]+(?:\.[0-9]+)?)(px|pt)(?:\s*\/\s*\S+)?\s+(.+)$/i.exec(css);
  if (!match) return undefined;
  const weightWord = (match[2] ?? "normal").toLowerCase();
  const weight = weightWord === "normal" ? 400 : weightWord === "bold" || weightWord === "bolder" ? 700 : weightWord === "lighter" ? 300 : Number(weightWord);
  const size = Number(match[3]) * (match[4].toLowerCase() === "pt" ? 4 / 3 : 1);
  if (!Number.isFinite(weight) || !(size > 0)) return undefined;
  return { stack: match[5].trim(), weight, italic: Boolean(match[1] && match[1].toLowerCase() !== "normal"), sizePx: size };
}

/** Why each character of a text resolves as it does (probes, debugging):
 * the face chosen in every family tried, its file's state and its check. */
export function explainText(text: string, font: EngineFont): unknown[] {
  refreshEpoch();
  const families = splitFamilyList(font.stack);
  return Array.from(text, (char) => {
    const codePoint = char.codePointAt(0)!;
    const steps: unknown[] = [];
    for (const family of families) {
      const faces = declaredFaces(family);
      const order = apiFaceOrder();
      const composite = selectComposite(faces, font.weight, font.italic, order === "pending" ? "last-wins" : order);
      const covering = composite.filter((face) => faceCovers(face, codePoint)).reverse();
      steps.push({
        family,
        declared: faces.length,
        composite: composite.length,
        covering: covering.slice(0, 3).map((face) => {
          const metrics = face.url ? faceMetrics(face.url) : "no-url";
          return {
            url: face.url,
            local: face.localNames.length,
            metrics: typeof metrics === "string" ? metrics : metrics.cmap.has(codePoint) ? "has-glyph" : "no-glyph",
            verified: faceVerification(face),
            detail: faceVerificationDetail(face),
          };
        }),
      });
      if (covering.length) break;
    }
    const resolved = resolveFace(families, font.weight, font.italic, codePoint);
    return { char, result: typeof resolved === "string" ? resolved : resolved.face.url, steps };
  });
}

/**
 * The height of one line box with `line-height: normal`, from the faces the
 * line's characters are drawn in and the paragraph's own primary face (the
 * strut), as CSS 2 §10.8 builds it: each face's ascent and descent (the
 * browsers round them to whole pixels) with half its line gap above and
 * below, the line spanning the highest top to the lowest bottom on a shared
 * baseline. Undefined unless every face is known.
 */
export function measureLineBoxHeight(runs: readonly { readonly text: string; readonly font: EngineFont }[], strut: EngineFont): number | undefined {
  const box = measureLineBox(runs, strut);
  return box === undefined ? undefined : box.top + box.bottom;
}

/** {@link measureLineBoxHeight}'s line box as its two halves: from the line's
 * top to the shared baseline, and from the baseline to its bottom. */
export function measureLineBox(runs: readonly { readonly text: string; readonly font: EngineFont }[], strut: EngineFont): { top: number; bottom: number } | undefined {
  refreshEpoch();
  let top = 0;
  let bottom = 0;
  const add = (face: FaceDeclaration, metrics: FaceMetrics, sizePx: number): void => {
    const unit = sizePx * face.sizeAdjust;
    const ascent = Math.round((face.ascentOverride ?? metrics.hhea.ascender / metrics.unitsPerEm) * unit);
    const descent = Math.round((face.descentOverride ?? -metrics.hhea.descender / metrics.unitsPerEm) * unit);
    const gap = Math.round((face.lineGapOverride ?? metrics.hhea.lineGap / metrics.unitsPerEm) * unit);
    top = Math.max(top, ascent + gap / 2);
    bottom = Math.max(bottom, descent + gap / 2);
  };
  // A wrapped column asks this of every character of every cell, over a few
  // fonts: 500 ms of a wheel run went to choosing the same faces again.
  const faceFinder = (font: EngineFont) => {
    const key = `${font.stack}\u0000${font.weight}\u0000${font.italic}\u0000${font.followLocalOnly}`;
    let faces = lineFaces.get(key);
    if (!faces) lineFaces.set(key, (faces = new Map()));
    const known = faces;
    const families = splitFamilyList(font.stack);
    return (codePoint: number) => {
      const held = known.get(codePoint);
      if (held) return held;
      const resolved = resolveFace(families, font.weight, font.italic, codePoint, font.followLocalOnly);
      if (typeof resolved !== "string") known.set(codePoint, resolved);
      return resolved;
    };
  };
  const strutFace = faceFinder(strut)(0x20);
  if (typeof strutFace === "string") return undefined;
  add(strutFace.face, strutFace.metrics, strut.sizePx);
  for (const run of runs) {
    const faceFor = faceFinder(run.font);
    const primary = faceFor(0x20);
    if (typeof primary === "string") return undefined;
    add(primary.face, primary.metrics, run.font.sizePx);
    const seen = new Set<string>();
    for (const grapheme of graphemesOf(run.text)) {
      let codePoint = grapheme.codePointAt(0)!;
      if (codePoint === 0xa0) codePoint = 0x20;
      if (codePoint < 0x20 || IGNORABLE.has(codePoint)) continue;
      const resolved = faceFor(codePoint);
      if (typeof resolved === "string") return undefined;
      const key = resolved.face.url ?? "";
      if (seen.has(key)) continue;
      seen.add(key);
      add(resolved.face, resolved.metrics, run.font.sizePx);
    }
  }
  return { top, bottom };
}
