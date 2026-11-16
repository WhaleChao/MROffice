/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { PDF_NATIVE_RULE_THICKNESS_EM, PDF_NATIVE_STRIKETHROUGH_OFFSET_EM, PDF_NATIVE_UNDERLINE_OFFSET_EM } from "./native-text-rule-range";
import type { PdfTextSourceAppearance } from "./source-text-appearance";
import { pdfTextDraftNodes, type PdfFallbackGlyphAdvances } from "./fallback-glyph-advances";

/**
 * The draft of a line being typed keeps the glyphs the reader did not touch
 * exactly where the page drew them.
 *
 * Between a keystroke and the exact page the browser's own layout of the
 * whole line was shown. Its face's advances are never the PDF's (kerned TJ
 * arrays, a wide word gap, a symbol set in another font), so the untouched
 * text jumped on every key and jumped back when the exact page arrived. The
 * page's last exact geometry (the source line, or the newest verified
 * candidate) gives every character its pen position: the characters before
 * and after the edit are set in boxes of exactly those advances, and only
 * what was typed in between is left to the browser.
 */
export interface PdfAnchoredDraftGlyph {
  readonly text: string;
  /** The glyph's advance in the geometry's own units: below zero where the
   * next glyph starts left of this one (the parts of a stack). */
  readonly advance: number;
  /** The glyph's size against the line's first glyph (a superscript set
   * smaller: 2/3); absent when it is the first glyph's size. */
  readonly scale?: number;
  /** The glyph's pen position below the first glyph's baseline, in the
   * geometry's own units (negative for a superscript, lifted); absent when
   * it sits on that baseline. */
  readonly rise?: number;
}

export interface PdfAnchoredDraftPlan {
  readonly prefix: readonly PdfAnchoredDraftGlyph[];
  readonly middle: string;
  readonly suffix: readonly PdfAnchoredDraftGlyph[];
}

type SourceCharacter = PdfTextSourceAppearance["characters"][number];
type Segment = Readonly<{ start: number; end: number; text: string; advance: number; scale?: number; rise?: number }>;










function segmentsOf(appearance: PdfTextSourceAppearance): Segment[] | null {
  const characters = appearance.characters;
  if (characters.length === 0) return null;
  // The row's measure in the geometry's own units: the glyph boxes' height
  // (a projected geometry keeps `fontSize` in the page's points while its
  // origins and quads are in CSS pixels; measured against the points, a
  // superscript's rise of 13 pt read as a second row at any zoom above
  
  const boxHeight = (character: SourceCharacter) =>
    Math.hypot(character.quad[4] - character.quad[0], character.quad[5] - character.quad[1]);
  const size = Math.max(...characters.map((character) => boxHeight(character) || Math.abs(character.fontSize) || 0));
  if (!(size > 0)) return null;
  const baseline = characters[0].origin[1];
  const firstSize = Math.abs(characters[0].fontSize) || size;
  const segments: Segment[] = [];
  let offset = 0;
  for (let index = 0; index < characters.length; index += 1) {
    const character = characters[index];
    const next = characters[index + 1];
    // A blank the box's text reads where no item draws one (a TeX word set off
    // by its position alone): no glyph of its own, the gap in the advance before it.
    if (character.start > offset && !appearance.text.slice(offset, character.start).trim()) {
      segments.push({ start: offset, end: character.start, text: appearance.text.slice(offset, character.start), advance: 0 });
      offset = character.start;
    }
    if (character.start !== offset || character.end <= character.start) return null;
    if (appearance.text.slice(character.start, character.end) !== character.text) return null;
    // One row: every pen position within nine tenths of a glyph of the first
    // one's baseline (a superscript is lifted a third of its size; a second row
    // lies a whole size or more away).
    const rise = character.origin[1] - baseline;
    // (A script of a script, a TeX x² in an exponent, stands three quarters of a
    // glyph box up; the next row, more than a whole one.)
    if (!Number.isFinite(rise) || Math.abs(rise) > size * 0.9) return null;
    const left = character.origin[0];
    const right = next ? next.origin[0] : Math.max(character.quad[2], character.quad[6]);
    const advance = right - left;
    // A pen that goes back is no line -- but for the part of a stack set
    // under or over the one before it (a fraction's denominator starting
    // left of where its numerator did: the next glyph stands on another
    // baseline, within a few glyphs back). Refused, one such row left the
    // whole block to the browser: the paragraph wrapped again in its frame
    // at the first key typed, and back when the page drew it (Attention
    
    const stacked = next !== undefined && Math.abs(next.origin[1] - character.origin[1]) > size * 0.1 &&
      advance >= -size * 3;
    if (!Number.isFinite(advance) || (advance < -size * 0.02 && !stacked)) return null;
    const own = Math.abs(character.fontSize) || firstSize;
    const scale = own / firstSize;
    segments.push({
      start: character.start, end: character.end, text: character.text,
      advance: advance < -size * 0.02 ? advance : Math.max(0, advance),
      ...(Math.abs(scale - 1) > 0.01 ? { scale: Math.round(scale * 1000) / 1000 } : {}),
      ...(Math.abs(rise) > size * 0.02 ? { rise: Math.round(rise * 1000) / 1000 } : {}),
    });
    offset = character.end;
  }
  return offset === appearance.text.length ? segments : null;
}

/**
 * Split `text` against the geometry's own text: the characters they share at
 * the start and at the end keep the geometry's advances; what lies between
 * is the reader's change. Null when the geometry cannot anchor a line (see
 * `segmentsOf`) or the text holds a line break.
 */
export function planPdfAnchoredTextDraft(
  anchor: PdfTextSourceAppearance | null | undefined,
  text: string,
): PdfAnchoredDraftPlan | null {
  if (!anchor || /[\r\n]/.test(text) || /[\r\n]/.test(anchor.text)) return null;
  const segments = segmentsOf(anchor);
  if (!segments) return null;
  const source = anchor.text;
  const limit = Math.min(source.length, text.length);
  let shared = 0;
  while (shared < limit && source[shared] === text[shared]) shared += 1;
  let tail = 0;
  while (tail < limit - shared && source[source.length - 1 - tail] === text[text.length - 1 - tail]) tail += 1;
  // Whole segments only: a surrogate pair or a ligature read as its letters
  // is anchored entire or left to the browser entire.
  const prefix = segments.filter((segment) => segment.end <= shared);
  const suffix = segments.filter((segment) => segment.start >= source.length - tail);
  const prefixEnd = prefix.length ? prefix[prefix.length - 1].end : 0;
  const suffixStart = suffix.length ? suffix[0].start : source.length;
  const middle = text.slice(prefixEnd, text.length - (source.length - suffixStart));
  const glyph = (segment: Segment): PdfAnchoredDraftGlyph => ({
    text: segment.text, advance: segment.advance,
    ...(segment.scale !== undefined ? { scale: segment.scale } : {}),
    ...(segment.rise !== undefined ? { rise: segment.rise } : {}),
  });
  return Object.freeze({
    prefix: Object.freeze(prefix.map(glyph)),
    middle,
    suffix: Object.freeze(suffix.map(glyph)),
  });
}

/**
 * Live Preview on a draft (components/office-ribbon/officeFormatPreview.tsx):
 * characters [start, end) of its text shown with `css` -- the font, size or
 * colour under the pointer -- until the pointer leaves.
 */
export interface PdfTextDraftFormatPreview {
  readonly start: number;
  readonly end: number;
  readonly css: Readonly<Record<string, string>>;
  /** They are the reader's selection: the selection band is drawn behind
   * them, at their previewed width. */
  readonly band: boolean;
  /** The rules the preview turns on or off, each on its own: an underline
   * previewed keeps a strikethrough the characters have, and the other way
   * round (composePdfTextDraftCss). */
  readonly decoration?: Readonly<{ underline?: boolean; strike?: boolean }>;
  /** The look moves no glyph (a colour, a rule, a bold drawn as a stroke of
   * the glyphs' own outlines): on a draft anchored to the page's glyphs the
   * previewed ones keep their boxes -- the page's own advances -- and wear
   * the look there. Set free for the browser to lay out, a line the PDF
   * sets tighter than its face's natural advances grew under the pointer
   * and came back when the press was written. */
  readonly keepsAdvances?: boolean;
  /** A look worn by a part of the line (a range, not a preview) that sets
   * the part's own glyphs again -- another face, at that face's advances:
   * on a draft anchored to the page's glyphs these characters leave their
   * boxes and are set by the browser in the look, the glyphs before them
   * where the page has them and the ones after them moved along as the
   * written line moves them. (Worn in their boxes, a part set in another
   * face drew that face's glyphs at the old face's advances.) */
  readonly frees?: boolean;
}









export function composePdfTextDraftCss(
  own: Readonly<Record<string, string>>,
  preview: PdfTextDraftFormatPreview,
): Record<string, string> {
  const css: Record<string, string> = { ...own, ...preview.css };
  const toggles = preview.decoration;
  if (toggles && (toggles.underline !== undefined || toggles.strike !== undefined)) {
    const lines = new Set((own["text-decoration-line"] ?? "").split(/\s+/u).filter((line) => line && line !== "none"));
    if (toggles.underline !== undefined) lines[toggles.underline ? "add" : "delete"]("underline");
    if (toggles.strike !== undefined) lines[toggles.strike ? "add" : "delete"]("line-through");
    css["text-decoration-line"] = lines.size ? [...lines].join(" ") : "none";
  } else if (own["text-decoration-line"] && !preview.css["text-decoration-line"]) {
    css["text-decoration-line"] = own["text-decoration-line"];
  }
  return css;
}

/** The look ranges give characters [from, to): every range over them, in order. */
function rangesCss(ranges: readonly PdfTextDraftFormatPreview[] | undefined, from: number, to: number): Record<string, string> {
  const css: Record<string, string> = {};
  for (const range of ranges ?? []) {
    if (range.start < to && range.end > from) Object.assign(css, range.css);
  }
  return css;
}

/** The selection band the PDF editor paints over a line's glyphs. */
const PREVIEW_SELECTION_BAND = "color-mix(in srgb, var(--console-accent-strong, #0f766e) 22%, transparent)";

function previewSpan(doc: Document, preview: PdfTextDraftFormatPreview): HTMLSpanElement {
  const span = doc.createElement("span");
  span.dataset.pdfTextDraftFormatPreview = "true";
  if (preview.band) span.style.backgroundColor = PREVIEW_SELECTION_BAND;
  for (const [property, value] of Object.entries(preview.css)) span.style.setProperty(property, value);
  return span;
}














export function pinPdfTextDraftLine(element: HTMLElement, largestPx: number, oneLine: boolean): void {
  const doc = element.ownerDocument;
  const layer = doc.createElement("div");
  layer.dataset.pdfTextDraftLayer = "true";
  layer.append(...element.childNodes);
  element.append(layer);
  if (oneLine) {
    const probe = () => {
      const span = doc.createElement("span");
      span.style.cssText = "display:inline-block;width:0;height:0;vertical-align:baseline";
      return span;
    };
    // The line's own: an empty line of the draft's font and line height.
    const strut = doc.createElement("div");
    strut.style.cssText = "position:absolute;left:0;top:0;visibility:hidden";
    const own = probe();
    strut.append(own);
    element.append(strut);
    const drawn = probe();
    layer.insertBefore(drawn, layer.firstChild);
    const shift = drawn.offsetTop - own.offsetTop;
    drawn.remove();
    strut.remove();
    if (Number.isFinite(shift) && Math.abs(shift) > 0.01) layer.style.transform = `translateY(${-Math.round(shift * 1000) / 1000}px)`;
  }
  const reach = Math.ceil(Math.max(0, largestPx) * 1.5);
  layer.style.paddingBlock = `${reach}px`;
  layer.style.marginBlock = `${-reach}px`;
  layer.style.willChange = "transform";
  if (oneLine) drawPdfTextDraftRules(layer);
}

/**
 * The rules of a line's characters drawn where the writer draws them: the
 * browser sets a line-through by its own face's metrics, 4 px off the writer's
 * at 20 px, and a press moved it. Each is a bar on its characters' own baseline,
 * in their size and colour, so it stands with them however large or raised.
 */
function drawPdfTextDraftRules(layer: HTMLElement): void {
  const doc = layer.ownerDocument;
  const view = doc.defaultView;
  if (!view) return;
  layer.style.position = "relative";
  const ruled = [...layer.querySelectorAll<HTMLElement>("span")].filter((span) => /underline|line-through/u.test(span.style.textDecorationLine));
  const bars: HTMLElement[] = [];
  for (const span of ruled) {
    const lines = span.style.textDecorationLine;
    const width = span.offsetWidth;
    if (!(width > 0)) continue;
    const look = view.getComputedStyle(span);
    const size = Number.parseFloat(look.fontSize);
    if (!(size > 0)) continue;
    const probe = doc.createElement("span");
    probe.style.cssText = "display:inline-block;width:0;height:0;vertical-align:baseline";
    span.insertBefore(probe, span.firstChild);
    const baseline = probe.offsetTop;
    const left = span.offsetLeft;
    probe.remove();
    span.style.textDecorationLine = "none";
    for (const [line, offset] of [["underline", PDF_NATIVE_UNDERLINE_OFFSET_EM], ["line-through", PDF_NATIVE_STRIKETHROUGH_OFFSET_EM]] as const) {
      if (!lines.includes(line)) continue;
      const bar = doc.createElement("div");
      bar.dataset.pdfTextDraftRule = line;
      const thick = size * PDF_NATIVE_RULE_THICKNESS_EM;
      bar.style.cssText = `position:absolute;pointer-events:none;left:${left}px;width:${width}px;` +
        `top:${baseline - (offset * size + thick)}px;height:${thick}px;background:${look.color}`;
      bars.push(bar);
    }
  }
  layer.append(...bars);
}

/** A draft laid out by the browser, with characters [start, end) previewed. */
export function fillPdfTextDraftPreview(
  element: HTMLElement,
  text: string,
  preview: PdfTextDraftFormatPreview,
  /** The PDF's advances of the characters the browser sets at none
   * (fallback-glyph-advances): each such one is a box of its advance. */
  glyphAdvances: PdfFallbackGlyphAdvances | null = null,
): void {
  const doc = element.ownerDocument;
  const start = Math.max(0, Math.min(preview.start, text.length));
  const end = Math.max(start, Math.min(preview.end, text.length));
  const span = previewSpan(doc, preview);
  span.replaceChildren(...pdfTextDraftNodes(doc, text.slice(start, end), glyphAdvances));
  element.replaceChildren(...pdfTextDraftNodes(doc, text.slice(0, start), glyphAdvances), span,
    ...pdfTextDraftNodes(doc, text.slice(end), glyphAdvances));
}

/** A draft laid out by the browser whose ranges are set in their own look:
 * the parts of a run set in a font the PDF does not hold yet, shown in that
 * font until the exact candidate is drawn. Ranges are ordered and disjoint. */
export function fillPdfTextDraftRanges(
  element: HTMLElement,
  text: string,
  ranges: readonly PdfTextDraftFormatPreview[],
  /** See fillPdfTextDraftPreview. */
  glyphAdvances: PdfFallbackGlyphAdvances | null = null,
): void {
  const doc = element.ownerDocument;
  const nodes: Node[] = [];
  let cursor = 0;
  for (const range of [...ranges].sort((left, right) => left.start - right.start)) {
    const start = Math.max(cursor, Math.min(range.start, text.length));
    const end = Math.max(start, Math.min(range.end, text.length));
    if (start > cursor) nodes.push(...pdfTextDraftNodes(doc, text.slice(cursor, start), glyphAdvances));
    if (end > start) {
      const span = previewSpan(doc, range);
      span.replaceChildren(...pdfTextDraftNodes(doc, text.slice(start, end), glyphAdvances));
      nodes.push(span);
    }
    cursor = end;
  }
  if (cursor < text.length) nodes.push(...pdfTextDraftNodes(doc, text.slice(cursor), glyphAdvances));
  element.replaceChildren(...nodes);
}








export function mergePdfTextDraftRanges(
  ranges: readonly PdfTextDraftFormatPreview[],
  preview: PdfTextDraftFormatPreview,
): readonly PdfTextDraftFormatPreview[] {
  const sorted = [...ranges].filter((range) => range.end > range.start).sort((left, right) => left.start - right.start);
  if (!(preview.end > preview.start)) return sorted;
  const merged: PdfTextDraftFormatPreview[] = [];
  let cursor = preview.start;
  for (const range of sorted) {
    if (range.end <= preview.start || range.start >= preview.end) {
      merged.push(range);
      continue;
    }
    // Before the preview, the range as it is; under it, both looks.
    if (range.start < preview.start) merged.push({ ...range, end: preview.start });
    const from = Math.max(range.start, preview.start), to = Math.min(range.end, preview.end);
    if (from > cursor) merged.push({ ...preview, start: cursor, end: from, css: composePdfTextDraftCss({}, preview) });
    merged.push({ start: from, end: to, css: composePdfTextDraftCss(range.css, preview), band: preview.band,
      ...(range.frees ? { frees: true } : {}) });
    cursor = to;
    if (range.end > preview.end) merged.push({ ...range, start: preview.end });
  }
  if (cursor < preview.end) merged.push({ ...preview, start: cursor, end: preview.end, css: composePdfTextDraftCss({}, preview) });
  return merged.sort((left, right) => left.start - right.start);
}

/**
 * Write the planned draft into `element`: anchored glyphs as inline boxes of
 * their own advance (`unitsToPx` turns the geometry's units into the
 * element's CSS pixels), the change between them as plain text. With a
 * preview, the characters it covers leave their boxes and are set in its
 * look by the browser; the glyphs around them keep their places.
 */
export function fillPdfAnchoredTextDraft(
  element: HTMLElement,
  plan: PdfAnchoredDraftPlan,
  unitsToPx: number,
  preview?: PdfTextDraftFormatPreview | null,
  /** Parts already set in a look of their own (a served family, a part
   * made a superscript before its row is written): each glyph box in such
   * a part wears the part's CSS, under the preview's. */
  ranges?: readonly PdfTextDraftFormatPreview[],
): void {
  const doc = element.ownerDocument;
  const rangeLook = (span: HTMLSpanElement, from: number, to: number) => {
    for (const range of ranges ?? []) {
      if (range.start < to && range.end > from) {
        for (const [property, value] of Object.entries(range.css)) span.style.setProperty(property, value);
        // A part set smaller keeps its glyphs end to end once written: its
        // boxes shrink with the size (left at the page's advance, "GE" at
        // 2/3 stood as "G E").
        const size = /^([\d.]+)%$/.exec(range.css["font-size"] ?? "");
        if (size && span.style.width) {
          span.style.width = `${Math.round(parseFloat(span.style.width) * Number(size[1]) / 100 * 1000) / 1000}px`;
        }
      }
    }
  };
  // A glyph's own size and rise on the row (a superscript: smaller,
  // lifted), kept whether it sits in its box or in a preview.
  const ownLook = (span: HTMLSpanElement, glyph: PdfAnchoredDraftGlyph) => {
    if (glyph.scale !== undefined) span.style.fontSize = `${Math.round(glyph.scale * 10000) / 100}%`;
    if (glyph.rise !== undefined) {
      span.style.position = "relative";
      span.style.top = `${Math.round(glyph.rise * unitsToPx * 1000) / 1000}px`;
    }
  };
  // Characters a range sets again in another face (`frees`): out of their
  // boxes, in one span a stretch of the same look.
  const freesAt = (from: number, to: number) => (ranges ?? []).some((range) => range.frees && range.start < to && range.end > from);
  let freed: { span: HTMLSpanElement; look: string } | null = null;
  const free = (into: Node[], text: string, from: number, glyph?: PdfAnchoredDraftGlyph) => {
    const css = rangesCss(ranges, from, from + text.length);
    const look = JSON.stringify([css, glyph?.scale ?? null, glyph?.rise ?? null]);
    if (!freed || freed.look !== look || into[into.length - 1] !== freed.span) {
      const span = doc.createElement("span");
      span.dataset.pdfAnchoredFreed = "true";
      // Like a glyph's box, at its face's own advances: the draft's letter
      // and word spacing are what fits a line the browser lays out whole to
      // the PDF's width, and are not this part's (inherited, they moved the
      
      span.style.letterSpacing = "0px";
      span.style.wordSpacing = "0px";
      for (const [property, value] of Object.entries(css)) span.style.setProperty(property, value);
      if (glyph) ownLook(span, glyph);
      freed = { span, look };
      into.push(span);
    }
    freed.span.append(text);
  };
  const box = (glyph: PdfAnchoredDraftGlyph, from: number) => {
    const span = doc.createElement("span");
    span.dataset.pdfAnchoredGlyph = "true";
    span.style.display = "inline-block";
    // (A pen that goes back -- the part of a stack -- is a box of no width
    // the next one is set back from.)
    if (glyph.advance < 0) {
      span.style.width = "0px";
      span.style.marginRight = `${Math.round(glyph.advance * unitsToPx * 1000) / 1000}px`;
    } else span.style.width = `${Math.round(glyph.advance * unitsToPx * 1000) / 1000}px`;
    span.style.whiteSpace = "pre";
    // The box is the PDF's advance; the browser's own spacing is not added.
    span.style.letterSpacing = "0px";
    span.style.wordSpacing = "0px";
    span.textContent = glyph.text;
    ownLook(span, glyph);
    rangeLook(span, from, from + glyph.text.length);
    // White space draws nothing, and its box is the PDF's advance: drawn in
    // a face the browser always has a space in. In the box's own face it
    // could be a page subset that maps the space to .notdef, which drew a
    // box between two words (IPCC title, 「MATE CHANGE」 after MATE was set
    
    if (/^\s+$/u.test(glyph.text)) span.style.fontFamily = "sans-serif";
    return span;
  };
  if (preview && preview.end > preview.start) {
    const nodes: Node[] = [];
    let span: HTMLSpanElement | null = null;
    const previewed = (text: string, from: number, glyph?: PdfAnchoredDraftGlyph) => {
      if (!span) {
        span = previewSpan(doc, preview);
        // (At the previewed face's own advances, as a freed part is.)
        if (!preview.css["letter-spacing"]) span.style.letterSpacing = "0px";
        if (!preview.css["word-spacing"]) span.style.wordSpacing = "0px";
        nodes.push(span);
      }
      // The characters keep their own look (a part's face, its rules)
      // under the preview's: composed, not replaced.
      const own = rangesCss(ranges, from, from + text.length);
      if (Object.keys(own).length || (glyph && (glyph.scale !== undefined || glyph.rise !== undefined))) {
        const inner = doc.createElement("span");
        inner.textContent = text;
        if (Object.keys(own).length) {
          for (const [property, value] of Object.entries(composePdfTextDraftCss(own, preview))) {
            inner.style.setProperty(property, value);
          }
        }
        if (glyph) ownLook(inner, glyph);
        span.append(inner);
      } else {
        span.append(text);
      }
    };
    // A glyph that keeps its box under the preview (keepsAdvances): its own
    // look and its part's, then the preview's over them, the band behind it.
    const previewedBox = (item: PdfAnchoredDraftGlyph, from: number) => {
      const boxed = box(item, from);
      boxed.dataset.pdfTextDraftFormatPreview = "true";
      if (preview.band) boxed.style.backgroundColor = PREVIEW_SELECTION_BAND;
      const own = rangesCss(ranges, from, from + item.text.length);
      for (const [property, value] of Object.entries(composePdfTextDraftCss(own, preview))) boxed.style.setProperty(property, value);
      // (Text typed after it starts a span of its own, after this box.)
      span = null;
      return boxed;
    };
    let offset = 0;
    const glyph = (item: PdfAnchoredDraftGlyph) => {
      const from = offset;
      offset += item.text.length;
      const inPreview = from < preview.end && offset > preview.start;
      if (!inPreview && freesAt(from, offset)) free(nodes, item.text, from, item);
      else if (!inPreview) nodes.push(box(item, from));
      else if (preview.keepsAdvances && !freesAt(from, offset)) nodes.push(previewedBox(item, from));
      else previewed(item.text, from, item);
    };
    plan.prefix.forEach(glyph);
    if (plan.middle) {
      const from = offset;
      const length = plan.middle.length;
      offset += length;
      const cut = Math.max(0, Math.min(preview.start - from, length));
      const until = Math.max(cut, Math.min(preview.end - from, length));
      if (cut > 0) nodes.push(doc.createTextNode(plan.middle.slice(0, cut)));
      if (until > cut) previewed(plan.middle.slice(cut, until), from + cut);
      if (until < length) nodes.push(doc.createTextNode(plan.middle.slice(until)));
    }
    plan.suffix.forEach(glyph);
    element.replaceChildren(...(nodes.length ? nodes : [doc.createTextNode("")]));
    return;
  }
  let at = 0;
  const nodes: Node[] = [];
  const set = (glyph: PdfAnchoredDraftGlyph) => {
    if (freesAt(at, at + glyph.text.length)) free(nodes, glyph.text, at, glyph);
    else nodes.push(box(glyph, at));
    at += glyph.text.length;
  };
  plan.prefix.forEach(set);
  if (plan.middle) {
    // What was typed wears the look of the run it was typed into.
    const own = rangesCss(ranges, at, at + plan.middle.length);
    if (Object.keys(own).length) {
      const span = doc.createElement("span");
      for (const [property, value] of Object.entries(own)) span.style.setProperty(property, value);
      span.textContent = plan.middle;
      nodes.push(span);
    } else nodes.push(doc.createTextNode(plan.middle));
    at += plan.middle.length;
  }
  plan.suffix.forEach(set);
  if (nodes.length === 0) nodes.push(doc.createTextNode(""));
  element.replaceChildren(...nodes);
}

/** One run of a draft set in runs: its characters and the CSS they wear. */
export interface PdfTextDraftRun {
  readonly text: string;
  readonly css: Readonly<Record<string, string>>;
}

/**
 * A draft set in runs (a text box with part of its text formatted): each
 * run a span in its own look, the text exactly the draft's (the caret and
 * selection read it character by character). With a preview, the
 * characters it covers wear its look over their run's.
 */
export function fillPdfTextDraftRuns(
  element: HTMLElement,
  runs: readonly PdfTextDraftRun[],
  preview?: PdfTextDraftFormatPreview | null,
): void {
  const doc = element.ownerDocument;
  const nodes: Node[] = [];
  let offset = 0;
  for (const run of runs) {
    const from = offset;
    const to = offset + run.text.length;
    offset = to;
    const cuts = [from, to];
    if (preview && preview.end > preview.start) {
      for (const cut of [preview.start, preview.end]) if (cut > from && cut < to) cuts.push(cut);
    }
    cuts.sort((a, b) => a - b);
    for (let index = 0; index + 1 < cuts.length; index += 1) {
      const start = cuts[index];
      const end = cuts[index + 1];
      if (end <= start) continue;
      const previewed = Boolean(preview && start >= preview.start && end <= preview.end && preview.end > preview.start);
      const span = previewed && preview ? previewSpan(doc, preview) : doc.createElement("span");
      span.dataset.pdfTextDraftRun = "true";
      for (const [property, value] of Object.entries(previewed && preview ? composePdfTextDraftCss(run.css, preview) : run.css)) {
        span.style.setProperty(property, value);
      }
      span.textContent = run.text.slice(start - from, end - from);
      nodes.push(span);
    }
  }
  element.replaceChildren(...(nodes.length ? nodes : [doc.createTextNode("")]));
}

/** One row of a block of several rows drafted on the page's own glyphs. */
export interface PdfAnchoredDraftRow {
  /** The row's first character's offset in the text drafted. */
  readonly start: number;
  /** The pen origin of the row's first glyph, in the geometry's units. */
  readonly origin: readonly [number, number];
  readonly plan: PdfAnchoredDraftPlan;
  /** What the text holds between this row and the next (the space the
   * reader joined the rows with, or nothing): it draws nothing. */
  readonly separator: string;
  /** The row the reader's change is in. */
  readonly edited: boolean;
}
















export function planPdfAnchoredTextRows(
  anchor: PdfTextSourceAppearance | null | undefined,
  text: string,
): readonly PdfAnchoredDraftRow[] | null {
  if (!anchor || /[\r\n]/.test(text) || /[\r\n]/.test(anchor.text)) return null;
  const characters = anchor.characters;
  if (characters.length < 2) return null;
  const boxHeight = (character: SourceCharacter) =>
    Math.hypot(character.quad[4] - character.quad[0], character.quad[5] - character.quad[1]);
  const size = Math.max(...characters.map((character) => boxHeight(character) || Math.abs(character.fontSize) || 0));
  if (!(size > 0)) return null;
  const groups: SourceCharacter[][] = [];
  let baseline = Number.NaN;
  for (const [index, character] of characters.entries()) {
    if (index && character.start < characters[index - 1].end) return null;
    // A new row: a pen position a glyph or more off the row's baseline
    // (segmentsOf keeps a superscript, lifted less, on its row).
    if (!groups.length || Math.abs(character.origin[1] - baseline) > size * 0.9) {
      groups.push([character]);
      baseline = character.origin[1];
    } else groups[groups.length - 1].push(character);
  }
  if (groups.length < 2) return null;
  const source = anchor.text;
  const spans = groups.map((group) => ({ start: group[0].start, end: group[group.length - 1].end, group }));
  if (spans[0].start !== 0 || spans[spans.length - 1].end !== source.length) return null;
  for (let index = 1; index < spans.length; index += 1) {
    if (source.slice(spans[index - 1].end, spans[index].start).trim()) return null;
    // Rows read top to bottom.
    if (!(spans[index].group[0].origin[1] > spans[index - 1].group[0].origin[1])) return null;
  }
  const limit = Math.min(source.length, text.length);
  let shared = 0;
  while (shared < limit && source[shared] === text[shared]) shared += 1;
  let tail = 0;
  while (tail < limit - shared && source[source.length - 1 - tail] === text[text.length - 1 - tail]) tail += 1;
  const changedEnd = source.length - tail;
  const edited = source === text ? -1 : spans.findIndex((span) => shared <= span.end);
  if (edited >= 0 && (shared < spans[edited].start || changedEnd > spans[edited].end)) return null;
  const grown = text.length - source.length;
  const rows: PdfAnchoredDraftRow[] = [];
  for (const [index, span] of spans.entries()) {
    const own = source.slice(span.start, span.end);
    const typed = index === edited
      ? source.slice(span.start, shared) + text.slice(shared, text.length - tail) + source.slice(changedEnd, span.end)
      : own;
    const row: PdfTextSourceAppearance = {
      ...anchor,
      text: own,
      characters: span.group.map((character) => ({ ...character, start: character.start - span.start, end: character.end - span.start })),
    };
    const plan = planPdfAnchoredTextDraft(row, typed);
    if (!plan) return null;
    const next = spans[index + 1];
    rows.push(Object.freeze({
      start: edited >= 0 && index > edited ? span.start + grown : span.start,
      origin: Object.freeze([span.group[0].origin[0], span.group[0].origin[1]]) as readonly [number, number],
      plan,
      separator: next ? source.slice(span.end, next.start) : "",
      edited: index === edited,
    }));
  }
  const drafted = rows.map((row) => row.plan.prefix.map((glyph) => glyph.text).join("") + row.plan.middle +
    row.plan.suffix.map((glyph) => glyph.text).join("") + row.separator).join("");
  return drafted === text ? Object.freeze(rows) : null;
}

/**
 * Fill a draft with its rows (planPdfAnchoredTextRows), each row set on its
 * own baseline at its first glyph's pen position. The draft's own transform
 * (a scale, a shift; never a turn or a slant) is read from the element, the
 * page's from its host: the rows stand where the page draws them whatever
 * the draft is laid out in. `breaks`: the text shows a line break between
 * rows (the input's rows) instead of the separator. False when the draft
 * cannot be placed so (no layout, a turned or slanted draft).
 */
export function fillPdfAnchoredTextRows(
  element: HTMLElement,
  rows: readonly PdfAnchoredDraftRow[],
  frame: Readonly<{ host: HTMLElement; pageWidth: number; pageHeight: number; breaks?: boolean;
    /** The page's other text (client rects): the edited row wraps before
     * the first of it to its right, as the page sets the paragraph again
     * in the room free beside it. */
    obstacles?: readonly Readonly<{ left: number; top: number; right: number; bottom: number }>[];
    /** The block is set centred (its rows on one axis): its left and right
     * edges, in the rows' own units. The row typed onto keeps the block's
     * axis, and wraps at the block's width, as the page sets it again. */
    centred?: Readonly<{ left: number; right: number }> | null }>,
  preview?: PdfTextDraftFormatPreview | null,
  ranges?: readonly PdfTextDraftFormatPreview[],
): boolean {
  const view = element.ownerDocument.defaultView;
  if (!view || !(frame.pageWidth > 0) || !(frame.pageHeight > 0)) return false;
  const host = frame.host.getBoundingClientRect();
  if (!(host.width > 0 && host.height > 0)) return false;
  const transform = view.getComputedStyle(element).transform;
  let a = 1, d = 1;
  if (transform && transform !== "none") {
    const values = /^matrix\(([^)]+)\)$/u.exec(transform)?.[1]?.split(",").map(Number);
    if (!values || values.length !== 6 || !values.every(Number.isFinite) ||
      Math.abs(values[1]) > 1e-6 || Math.abs(values[2]) > 1e-6 || !(values[0] > 0) || !(values[3] > 0)) return false;
    [a, , , d] = values;
  }
  const doc = element.ownerDocument;
  const shifted = (range: PdfTextDraftFormatPreview, by: number): PdfTextDraftFormatPreview =>
    ({ ...range, start: range.start - by, end: range.end - by });
  const layers: HTMLDivElement[] = [];
  for (const [index, row] of rows.entries()) {
    const layer = doc.createElement("div");
    layer.dataset.pdfAnchoredRow = String(index);
    layer.style.cssText = "position:absolute;left:0;top:0;white-space:pre;margin:0;padding:0";
    const length = row.plan.prefix.reduce((sum, glyph) => sum + glyph.text.length, 0) + row.plan.middle.length +
      row.plan.suffix.reduce((sum, glyph) => sum + glyph.text.length, 0);
    const own = preview && preview.start < row.start + length && preview.end > row.start ? shifted(preview, row.start) : null;
    const looks = ranges?.filter((range) => range.start < row.start + length && range.end > row.start)
      .map((range) => shifted(range, row.start));
    fillPdfAnchoredTextDraft(layer, row.plan, host.width / frame.pageWidth / a, own, looks);
    const separator = frame.breaks && index < rows.length - 1 ? "\n" : row.separator;
    if (separator) {
      const gap = doc.createElement("span");
      gap.dataset.pdfAnchoredRowSeparator = "true";
      gap.style.cssText = "display:inline-block;width:0;height:0;overflow:hidden;vertical-align:top;white-space:pre";
      gap.textContent = separator;
      layer.append(gap);
    }
    layers.push(layer);
  }
  element.replaceChildren(...layers);
  const corner = element.getBoundingClientRect();
  const baselineOf = (layer: HTMLElement) => {
    // The row's baseline in the layer: an empty box sat on it.
    const probe = doc.createElement("span");
    probe.style.cssText = "display:inline-block;width:0;height:0;vertical-align:baseline";
    layer.insertBefore(probe, layer.firstChild);
    const baseline = probe.offsetTop;
    probe.remove();
    return baseline;
  };
  const place = (layer: HTMLElement, x: number, y: number): boolean => {
    const baseline = baselineOf(layer);
    if (![x, y, baseline].every(Number.isFinite)) return false;
    layer.style.transform = `translate(${Math.round(x * 1000) / 1000}px, ${Math.round((y - baseline) * 1000) / 1000}px)`;
    return true;
  };
  // The rows' pitch on the page (client px), for rows a wrap adds.
  const baselines = rows.map((row) => host.top + row.origin[1] * host.height / frame.pageHeight);
  const pitch = baselines.length > 1 ? (baselines[baselines.length - 1] - baselines[0]) / (baselines.length - 1) : 0;
  let added = 0;
  for (const [index, row] of rows.entries()) {
    const layer = layers[index];
    const x = (host.left + row.origin[0] * host.width / frame.pageWidth - corner.left) / a;
    const y = (baselines[index] + added * pitch - corner.top) / d;
    if (!place(layer, x, y)) return false;
    if (!row.edited) continue;
    if (frame.centred) {
      // A centred block: the row keeps the block's axis as it is typed onto,
      // and what no longer fits the block's width goes on rows of its own
      // under it, on that axis too (the page sets the block so: IPCC SPM.1's
      
      // the axis until the page's own preview arrived).
      const axis = host.left + (frame.centred.left + frame.centred.right) / 2 * host.width / frame.pageWidth;
      const width = (frame.centred.right - frame.centred.left) * host.width / frame.pageWidth;
      const centre = (target: HTMLElement, line: number) => {
        // (Its ink: the separator after it draws nothing.)
        const range = doc.createRange();
        range.selectNodeContents(target);
        const rects = Array.from(range.getClientRects()).filter((rect) => rect.width > 0);
        range.detach();
        if (!rects.length) return true;
        const left = Math.min(...rects.map((rect) => rect.left)), right = Math.max(...rects.map((rect) => rect.right));
        const own = target.getBoundingClientRect().left;
        return place(target, (axis - (right - left) / 2 - (left - own) - corner.left) / a,
          (baselines[index] + line * pitch - corner.top) / d);
      };
      let current: HTMLElement = layer;
      for (let wraps = 0; wraps < 64 && pitch > 0; wraps += 1) {
        const band = current.getBoundingClientRect();
        const rest = splitPdfDraftRowAt(current, band.left + width + 0.5);
        if (!rest) break;
        current.after(rest);
        if (!centre(current, added)) return false;
        added += 1;
        if (!place(rest, x, (baselines[index] + added * pitch - corner.top) / d)) return false;
        current = rest;
      }
      if (!centre(current, added)) return false;
      continue;
    }
    if (!(pitch > 0)) continue;
    // Typed past the room beside it, the row's last words go on rows of
    // their own under it, from its left edge, as the page sets them.
    let current: HTMLElement = layer;
    for (let wraps = 0; wraps < 64; wraps += 1) {
      const band = current.getBoundingClientRect();
      const start = host.left + row.origin[0] * host.width / frame.pageWidth;
      const limit = Math.min(host.right - host.width * 0.02, ...(frame.obstacles ?? [])
        .filter((box) => box.left >= start - 1 && box.top < band.bottom && box.bottom > band.top)
        .map((box) => box.left));
      const rest = splitPdfDraftRowAt(current, limit);
      if (!rest) break;
      current.after(rest);
      added += 1;
      if (!place(rest, x, (baselines[index] + added * pitch - corner.top) / d)) return false;
      current = rest;
    }
  }
  return true;
}

/** Split a drafted row before the first word that ends past `limit` (client
 * x): the words from there on, in a layer of their own (the caller places
 * it), the space they followed left at the row's end. Null when the row
 * fits, or its first word alone is wider than the room. */
function splitPdfDraftRowAt(layer: HTMLElement, limit: number): HTMLElement | null {
  const doc = layer.ownerDocument;
  const content = layer.getBoundingClientRect();
  if (!(content.right > limit + 0.5)) return null;
  // Every place after a blank (in a text, or a blank's own box), in order.
  const breaks: Array<{ node: Node; offset: number }> = [];
  const walk = (node: Node) => {
    if (node.nodeType === Node.TEXT_NODE) {
      const text = node.textContent ?? "";
      for (let index = 0; index < text.length; index += 1) {
        if (/\s/u.test(text[index]) && !/\s/u.test(text[index + 1] ?? "x")) breaks.push({ node, offset: index + 1 });
      }
      return;
    }
    if (node instanceof HTMLElement && node.dataset.pdfAnchoredRowSeparator === "true") return;
    for (const child of Array.from(node.childNodes)) walk(child);
  };
  walk(layer);
  const range = doc.createRange();
  for (let index = breaks.length - 1; index >= 0; index -= 1) {
    const at = breaks[index];
    // The words before the break, without the blank it is.
    range.setStart(layer, 0);
    range.setEnd(at.node, at.offset - 1);
    const rects = Array.from(range.getClientRects()).filter((rect) => rect.width > 0);
    if (!rects.length) return null;
    if (Math.max(...rects.map((rect) => rect.right)) > limit + 0.5) continue;
    // The rest: from the break to the row's end, moved whole. A row is
    // shallow: its glyph boxes, the typed text (bare, or in a span of its
    // run's look), the separator.
    const text = at.node as Text;
    const rest = doc.createElement("div");
    rest.dataset.pdfAnchoredRow = `${layer.dataset.pdfAnchoredRow ?? ""}+`;
    rest.style.cssText = layer.style.cssText;
    const owner = text.parentNode === layer ? text : text.parentNode as HTMLElement;
    if (owner.parentNode !== layer) return null;
    const after: Node[] = [];
    for (let sibling = owner.nextSibling; sibling; sibling = sibling.nextSibling) after.push(sibling);
    const moved: Node[] = [];
    if (at.offset < (text.textContent ?? "").length) {
      const tail = text.splitText(at.offset);
      if (owner === text) moved.push(tail);
      else {
        const copy = (owner as HTMLElement).cloneNode(false);
        copy.appendChild(tail);
        moved.push(copy);
      }
    }
    rest.append(...moved, ...after);
    range.detach();
    return rest;
  }
  range.detach();
  return null;
}
