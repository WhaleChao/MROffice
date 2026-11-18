/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure deck-wide FONT REPLACE (PowerPoint Home ▸ Replace ▸ Replace Fonts): swap
 * every use of one font family for another across all slides. A font lives in
 * three places — a block's `style.fontFamily`, a per-run `run.style.fontFamily`,
 * and (for table blocks) each cell's `style` / `runs`. Operates structurally on
 * the heterogeneous slide model so it is decoupled from the concrete block types,
 * IMMUTABLE (returns new objects only where something changed; unchanged blocks /
 * slides are returned by reference), and unit-testable apart from the editor.
 */

import { primaryOfficeFontFace } from "../components/office-ribbon/officeFontFace";
import {
  isThemeFontReference,
  MASTER_FONT_LEVEL_KEYS,
  MASTER_FONT_SCRIPTS,
  MASTER_TEXT_STYLE_NAMES,
  type SlideMaster,
} from "./presentation-master";






















export interface FontReplacement {
  readonly id: string;
  readonly from: string;
  readonly to: string;
}

/** The deck's log holds at most this many replacements (the server's bound). */
export const FONT_REPLACEMENTS_MAX = 200;
const FONT_REPLACEMENT_ID_MAX = 64;
const FONT_REPLACEMENT_FACE_MAX = 120;

/** A deck's `fontReplacements` as it may be kept: entries with an id, a face
 * replaced and its replacement (one family each, never a theme reference),
 * the first of each id. Anything else is dropped. */
export function normalizeFontReplacements(raw: unknown): FontReplacement[] {
  if (!Array.isArray(raw)) return [];
  const out: FontReplacement[] = [];
  const seen = new Set<string>();
  for (const item of raw) {
    if (out.length >= FONT_REPLACEMENTS_MAX) break;
    if (!item || typeof item !== "object") continue;
    const { id, from, to } = item as Record<string, unknown>;
    if (typeof id !== "string" || !id || id.length > FONT_REPLACEMENT_ID_MAX || seen.has(id)) continue;
    if (typeof from !== "string" || typeof to !== "string") continue;
    const source = from.trim();
    const target = to.trim();
    if (!source || !target || source.length > FONT_REPLACEMENT_FACE_MAX || target.length > FONT_REPLACEMENT_FACE_MAX) continue;
    if (isThemeFontReference(source) || isThemeFontReference(target)) continue;
    seen.add(id);
    out.push({ id, from: source, to: target });
  }
  return out;
}

/** `{face (lower case): the face it ends as}` for a whole log, in order: each
 * entry moves every face that, after the entries before it, reads as its
 * `from` (the server's `compose_font_replacements`). */
export function composeFontReplacements(log: readonly FontReplacement[]): Map<string, string> {
  const current = new Map<string, string>();
  for (const { from, to } of log) {
    const folded = from.trim().toLowerCase();
    if (!folded || folded === to.trim().toLowerCase()) continue;
    for (const [key, value] of current) {
      if (value.toLowerCase() === folded) current.set(key, to.trim());
    }
    if (!current.has(folded)) current.set(folded, to.trim());
  }
  for (const [key, value] of current) if (key === value.toLowerCase()) current.delete(key);
  return current;
}

/** The master / layout placeholders whose words are their text styles' sample
 * (the import's `placeholder` for title / ctrTitle, subTitle, body / obj). */
const STYLE_SAMPLE_PLACEHOLDERS = new Set(["title", "subtitle", "content"]);

function isStyleSample(block: unknown): boolean {
  const placeholder = block && typeof block === "object" ? (block as { placeholder?: unknown }).placeholder : undefined;
  return typeof placeholder === "string" && STYLE_SAMPLE_PLACEHOLDERS.has(placeholder);
}
import { isNativeSmartArtMember } from "./presentation-native-smartart";
import { tableCellTypefaceFromChoice } from "./presentation-table-style";

function styleFont(style: unknown): string | undefined {
  if (style && typeof style === "object") {
    const f = (style as { fontFamily?: unknown }).fontFamily;
    if (typeof f === "string" && f) return f;
  }
  return undefined;
}

/**
 * One font a master or its theme provides, for the Replace list (§73).
 *
 * `editable` is true for a face the master's own `p:txStyles` names: Replace
 * Fonts changes it on every level that states it. A theme font (what a
 * `+mj-lt` / `+mn-lt` reference resolves to) is listed so the deck's fonts
 * are all there, but it is the THEME's -- PowerPoint changes it under Slide
 * Master ▸ Fonts / Design ▸ Fonts, and Microsoft's Replace Fonts page says
 * nothing about the theme, so it is not replaced here.
 */
export interface MasterFontEntry {
  face: string;
  editable: boolean;
}

/**
 * Collect the fonts the masters provide: literal faces from each master's
 * `textStyleFonts` (editable) and the theme faces its references resolve to
 * (read-only). Sorted, de-duped; a face that is both is editable.
 */
export function collectMasterFonts(
  masters: readonly SlideMaster[] | undefined,
): MasterFontEntry[] {
  const editable = new Set<string>();
  const themed = new Set<string>();
  for (const master of masters ?? []) {
    for (const name of MASTER_TEXT_STYLE_NAMES) {
      const faces = master.textStyleFonts?.[name];
      // Level one, then every deeper level that states a face of its own:
      // a body style whose level four is Consolas offers Consolas here too
      // (§74). Before, the list showed level one's face and nothing else.
      const slots = [
        faces,
        ...MASTER_FONT_LEVEL_KEYS.map((key) => faces?.levels?.[key]),
      ];
      for (const level of slots) {
        for (const script of MASTER_FONT_SCRIPTS) {
          const face = level?.[script];
          if (!face) continue;
          if (isThemeFontReference(face)) {
            const family = face.startsWith("+mj") ? "major" : "minor";
            const resolved = master.themeFonts?.[family]?.[script];
            if (resolved) themed.add(resolved);
          } else {
            editable.add(face);
          }
        }
      }
    }
  }
  const entries = [...editable].map((face) => ({ face, editable: true }));
  for (const face of themed) {
    if (!editable.has(face)) entries.push({ face, editable: false });
  }
  return entries.sort((a, b) => a.face.localeCompare(b.face));
}

/** A font as Replace Fonts names and matches it: its primary face, the
 *  font a run draws in, whatever fallbacks the entry that set it listed
 *  (Home writes "Arial, Helvetica, sans-serif", the advanced format another
 *  stack; both are Arial, PPT-B301). Compared without regard to case, as
 *  Office compares font names. */
function fontName(value: unknown): string {
  // A value naming no face (a bare generic family such as "serif") is its
  // own name, so it is still listed and replaced.
  return primaryOfficeFontFace(value) || (typeof value === "string" ? value.trim() : "");
}

function fontIdentity(value: unknown): string {
  return fontName(value).toLowerCase();
}

/** Collect every font in use across the deck, by its primary face (sorted,
 *  de-duped): the names Replace Fonts lists. */
export function collectDeckFonts<S extends { blocks?: unknown[] }>(
  slides: readonly S[],
): string[] {
  const fonts = new Map<string, string>();
  const add = (value: string | undefined) => {
    const face = fontName(value);
    if (face && !fonts.has(face.toLowerCase())) fonts.set(face.toLowerCase(), face);
  };
  const visit = (carrier: unknown) => {
    if (!carrier || typeof carrier !== "object") return;
    const rec = carrier as Record<string, unknown>;
    add(styleFont(rec.style));
    // A table cell's preset font (Table Design, or Home > Font on an empty
    // cell) is the cell's own `fontFamily` (PPT-B379).
    if (typeof rec.fontFamily === "string") add(rec.fontFamily);
    if (Array.isArray(rec.runs)) {
      for (const run of rec.runs) add(styleFont((run as { style?: unknown })?.style));
    }
  };
  for (const slide of slides) {
    for (const block of slide.blocks ?? []) {
      visit(block);
      // A list's character formatting lives on its paragraphs' runs.
      const paragraphs = (block as { paragraphs?: unknown }).paragraphs;
      if (Array.isArray(paragraphs)) for (const paragraph of paragraphs) visit(paragraph);
      const rows = (block as { rows?: unknown }).rows;
      if (Array.isArray(rows)) {
        for (const row of rows) {
          if (Array.isArray(row)) for (const cell of row) visit(cell);
        }
      }
    }
  }
  return [...fonts.values()].sort();
}

function swapStyle(style: unknown, from: string, to: string, literalOnly = false): unknown {
  const font = styleFont(style);
  if (literalOnly && style && typeof style === "object" && (style as { themeFontRole?: unknown }).themeFontRole) {
    return style;
  }
  if (font && fontIdentity(font) === fontIdentity(from)) {
    // Text that follows the theme's font (+mn-lt / +mj-lt, or no face of its
    // own) and draws in `from` takes `to` as a face of its own: PowerPoint
    // for Windows writes it as an explicit <a:latin> and leaves the theme's
    
    // runs, the master's and layouts' field runs too). Left marked as the
    // theme's, the page showed the new face and the file never had it.
    const next: Record<string, unknown> = { ...(style as Record<string, unknown>), fontFamily: to };
    delete next.themeFontRole;
    if (Array.isArray(next.inheritedKeys)) {
      const kept = (next.inheritedKeys as unknown[]).filter((key) => key !== "fontFamily");
      if (kept.length) next.inheritedKeys = kept;
      else delete next.inheritedKeys;
    }
    return next;
  }
  return style;
}

/** Swap a text carrier's (block / cell) `style` + per-run fonts. Returns the
 *  input by reference when nothing changed. */
function swapCarrier(carrier: unknown, from: string, to: string, literalOnly = false): unknown {
  if (!carrier || typeof carrier !== "object") return carrier;
  const rec = carrier as Record<string, unknown>;
  let next = rec;
  const nextStyle = swapStyle(rec.style, from, to, literalOnly);
  if (nextStyle !== rec.style) next = { ...next, style: nextStyle };
  if (
    typeof rec.fontFamily === "string" &&
    fontIdentity(rec.fontFamily) === fontIdentity(from)
  ) {
    // A table cell's own font is a typeface (`a:latin@typeface`), not a
    // browser stack: the stack the font list hands over was stored whole and
    // drawn as one quoted, unknown family, so the cells fell back to the
    
    next = { ...next, fontFamily: tableCellTypefaceFromChoice(to) ?? to };
  }
  if (Array.isArray(rec.runs)) {
    let changed = false;
    const runs = rec.runs.map((run) => {
      if (run && typeof run === "object") {
        const ns = swapStyle((run as { style?: unknown }).style, from, to, literalOnly);
        if (ns !== (run as { style?: unknown }).style) {
          changed = true;
          return { ...(run as object), style: ns };
        }
      }
      return run;
    });
    if (changed) next = { ...next, runs };
  }
  return next;
}

/** One block with `from` replaced by `to`: its own style and runs, its list
 * paragraphs' runs, its table cells. `literalOnly` (a SmartArt member's
 * words): only a face the words name, not one that follows the theme.
 * Returned by reference when nothing changed. */
function swapBlock(block: unknown, from: string, to: string, literalOnly = false): unknown {
  let nextBlock = swapCarrier(block, from, to, literalOnly);
  // A list's character formatting lives on its paragraphs' runs, which
  // the canvas draws and saves (PPT-B085).
  const paragraphs = (block as { paragraphs?: unknown }).paragraphs;
  if (Array.isArray(paragraphs)) {
    let paragraphsChanged = false;
    const nextParagraphs = paragraphs.map((paragraph) => {
      const swapped = swapCarrier(paragraph, from, to, literalOnly);
      if (swapped !== paragraph) paragraphsChanged = true;
      return swapped;
    });
    if (paragraphsChanged) {
      nextBlock = { ...(nextBlock as object), paragraphs: nextParagraphs };
    }
  }
  const rows = (block as { rows?: unknown }).rows;
  if (Array.isArray(rows)) {
    let rowsChanged = false;
    const nextRows = rows.map((row) => {
      if (!Array.isArray(row)) return row;
      let rowChanged = false;
      const cells = row.map((cell) => {
        const nc = swapCarrier(cell, from, to, literalOnly);
        if (nc !== cell) rowChanged = true;
        return nc;
      });
      if (rowChanged) {
        rowsChanged = true;
        return cells;
      }
      return row;
    });
    if (rowsChanged) {
      nextBlock = { ...(nextBlock as object), rows: nextRows };
    }
  }
  return nextBlock;
}

function swapBlocks<B>(blocks: readonly B[], from: string, to: string, skip?: (block: B) => boolean): B[] | null {
  let changed = false;
  const next = blocks.map((block) => {
    if (skip?.(block)) return block;
    // A file's SmartArt: only the words that name the face outright move
    
    // words that follow the theme's body font did not).
    const swapped = swapBlock(block, from, to, isNativeSmartArtMember(block)) as B;
    if (swapped !== block) changed = true;
    return swapped;
  });
  return changed ? next : null;
}

/**
 * Return a new slide list with every `from` font replaced by `to` (block styles,
 * run styles, list paragraphs' runs, and table-cell styles / runs). No-op clone when `from` is empty or
 * equals `to`. Unchanged blocks / slides are kept by reference so a deck that
 * doesn't use `from` serializes byte-for-byte.
 */
export function replaceDeckFont<S extends { blocks?: unknown[] }>(
  slides: readonly S[],
  from: string,
  to: string,
): S[] {
  if (!from || from === to || !fontIdentity(from)) return slides.slice();
  return slides.map((slide) => {
    if (!Array.isArray(slide.blocks)) return slide;
    const blocks = swapBlocks(slide.blocks, from, to);
    return blocks ? ({ ...slide, blocks } as S) : slide;
  });
}

/**
 * The masters with `from` replaced by `to` on the words of their own and
 * their layouts' shapes the deck holds (`blocks`), as PowerPoint writes it:
 * the date, footer, number and picture placeholders and the other shapes move,
 * the title, subtitle and body placeholders (their text styles' sample) and
 * the text styles themselves (`textStyleFonts`) stay. Unchanged masters come
 * back by reference.
 */
export function replaceMasterShapeFonts(
  masters: readonly SlideMaster[],
  from: string,
  to: string,
): SlideMaster[] {
  if (!from || from === to || !fontIdentity(from)) return masters.slice();
  return masters.map((master) => {
    let next = master;
    if (Array.isArray(master.blocks)) {
      const blocks = swapBlocks(master.blocks, from, to, isStyleSample);
      if (blocks) next = { ...next, blocks };
    }
    if (Array.isArray(master.layouts)) {
      let layoutsChanged = false;
      const layouts = master.layouts.map((layout) => {
        const own = (layout as { blocks?: unknown[] }).blocks;
        if (!Array.isArray(own)) return layout;
        const blocks = swapBlocks(own, from, to, isStyleSample);
        if (!blocks) return layout;
        layoutsChanged = true;
        return { ...layout, blocks } as typeof layout;
      });
      if (layoutsChanged) next = { ...next, layouts };
    }
    return next;
  });
}

/** A card's blocks (a layout's or master's drawing the server sends beside the
 * page) as the deck's Replace Fonts log draws them. */
export function cardBlocksWithFontReplacements<B>(blocks: readonly B[], log: readonly FontReplacement[]): readonly B[] {
  if (log.length === 0) return blocks;
  let out: readonly B[] = blocks;
  for (const { from, to } of log) out = swapBlocks(out, from, to, isStyleSample) ?? out;
  return out;
}

/**
 * One Replace Fonts on a deck: its slides, its masters' and layouts' shapes,
 * and its log with the entry `id` appended. Null when nothing in the deck
 * draws in `from` (the dialog's list offers no such font) or `from` is `to`.
 * The log names each end by its one family (`a:latin/@typeface`), whatever
 * stack the font list handed over.
 */
export function replaceFontsInDeck<S extends { blocks?: unknown[] }>(
  deck: { slides: readonly S[]; masters: readonly SlideMaster[]; log: readonly FontReplacement[] },
  fromValue: string,
  toValue: string,
  id: string,
): { slides: S[]; masters: SlideMaster[]; log: FontReplacement[] } | null {
  const from = fontName(fromValue);
  const to = primaryOfficeFontFace(toValue) || toValue.trim();
  if (!from || !to || fontIdentity(from) === fontIdentity(to) || isThemeFontReference(from) || isThemeFontReference(to)) {
    return null;
  }
  const slides = replaceDeckFont(deck.slides, fromValue, toValue);
  const masters = replaceMasterShapeFonts(deck.masters, fromValue, toValue);
  const log = normalizeFontReplacements([...deck.log, { id, from, to }]);
  if (log.length === deck.log.length) return null;
  return { slides, masters, log };
}

/** A new entry's id: unique across the windows that may append at once. */
export function newFontReplacementId(): string {
  const random =
    typeof crypto !== "undefined" && typeof crypto.randomUUID === "function"
      ? crypto.randomUUID()
      : `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
  return `rf-${random}`;
}

/** A stored deck's log: its entries when it states the list (an empty one
 * included), undefined when it states none. */
export function deckFontReplacementsOf(content: unknown): FontReplacement[] | undefined {
  const raw = content && typeof content === "object" ? (content as { fontReplacements?: unknown }).fontReplacements : undefined;
  return Array.isArray(raw) ? normalizeFontReplacements(raw) : undefined;
}
