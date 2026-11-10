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

import { primaryOfficeFontFace } from "@mrnote/core/web/components/office-ribbon/officeFontFace.ts";
import {
  isThemeFontReference,
  MASTER_FONT_LEVEL_KEYS,
  MASTER_FONT_SCRIPTS,
  MASTER_TEXT_STYLE_NAMES,
  type SlideMaster,
} from "@mrnote/core/web/lib/presentation-master.ts";
import { tableCellTypefaceFromChoice } from "@mrnote/core/web/lib/presentation-table-style.ts";

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

/**
 * Return new masters with every literal `from` face in `textStyleFonts`
 * replaced by `to`. Theme references are never touched. Unchanged masters
 * are kept by reference so a deck whose masters do not use `from`
 * serializes byte-for-byte.
 *
 * `textStyleFonts` holds the OOXML typeface VERBATIM, and `a:latin/@typeface`
 * names one family; the dialog's lists are CSS stacks (a slide font) or bare
 * faces (a master font), so both ends are reduced to their primary family
 * here. Picking the deck's "Arial, 'Helvetica Neue', ..." therefore also
 * replaces the master's literal `Arial`, which is what "replace this font
 * throughout the presentation" means.
 */
export function replaceMasterFonts(
  masters: readonly SlideMaster[],
  fromValue: string,
  toValue: string,
): SlideMaster[] {
  const from = primaryOfficeFontFace(fromValue);
  const to = primaryOfficeFontFace(toValue);
  if (!from || !to || from === to || isThemeFontReference(from)) {
    return masters.slice();
  }
  return masters.map((master) => {
    if (!master.textStyleFonts) return master;
    let changed = false;
    const textStyleFonts: NonNullable<SlideMaster["textStyleFonts"]> = {};
    for (const name of MASTER_TEXT_STYLE_NAMES) {
      const faces = master.textStyleFonts[name];
      if (!faces) continue;
      let nextFaces = faces;
      for (const script of MASTER_FONT_SCRIPTS) {
        if (faces[script] === from) {
          nextFaces = { ...nextFaces, [script]: to };
          changed = true;
        }
      }
      // Each deeper level that states this face moves too, and a level that
      // states a DIFFERENT font keeps it -- that is the whole point of
      // projecting the nine levels (§74).
      if (faces.levels) {
        let nextLevels = faces.levels;
        for (const key of MASTER_FONT_LEVEL_KEYS) {
          const level = faces.levels[key];
          if (!level) continue;
          let nextLevel = level;
          for (const script of MASTER_FONT_SCRIPTS) {
            if (level[script] === from) {
              nextLevel = { ...nextLevel, [script]: to };
            }
          }
          if (nextLevel !== level) {
            nextLevels = { ...nextLevels, [key]: nextLevel };
            changed = true;
          }
        }
        if (nextLevels !== faces.levels) nextFaces = { ...nextFaces, levels: nextLevels };
      }
      textStyleFonts[name] = nextFaces;
    }
    return changed ? { ...master, textStyleFonts } : master;
  });
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

function swapStyle(style: unknown, from: string, to: string): unknown {
  const font = styleFont(style);
  if (font && fontIdentity(font) === fontIdentity(from)) {
    return { ...(style as object), fontFamily: to };
  }
  return style;
}

/** Swap a text carrier's (block / cell) `style` + per-run fonts. Returns the
 *  input by reference when nothing changed. */
function swapCarrier(carrier: unknown, from: string, to: string): unknown {
  if (!carrier || typeof carrier !== "object") return carrier;
  const rec = carrier as Record<string, unknown>;
  let next = rec;
  const nextStyle = swapStyle(rec.style, from, to);
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
        const ns = swapStyle((run as { style?: unknown }).style, from, to);
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
    let slideChanged = false;
    const blocks = slide.blocks.map((block) => {
      let nextBlock = swapCarrier(block, from, to);
      // A list's character formatting lives on its paragraphs' runs, which
      // the canvas draws and saves (PPT-B085).
      const paragraphs = (block as { paragraphs?: unknown }).paragraphs;
      if (Array.isArray(paragraphs)) {
        let paragraphsChanged = false;
        const nextParagraphs = paragraphs.map((paragraph) => {
          const swapped = swapCarrier(paragraph, from, to);
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
            const nc = swapCarrier(cell, from, to);
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
      if (nextBlock !== block) slideChanged = true;
      return nextBlock;
    });
    return slideChanged ? ({ ...slide, blocks } as S) : slide;
  });
}

