/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Insert Symbol for presentation slide text (PowerPoint / WPS "Insert → Symbol"
 * parity), ported from the native-document picker.
 *
 * The CATALOG of glyphs is shared verbatim with the document editor
 * (`native-document/insertSymbols`) so the two editors offer the exact same
 * curated set (legal/typographic marks, math operators, arrows, currency, a few
 * decorative marks). This module re-exports that catalog and adds ONE tiny,
 * pure caret-splice used by the presentation text toolbar — the slide text is a
 * flat string (not a TipTap document), so inserting a glyph is a plain string
 * splice at the caret rather than an editor command.
 *
 * Keep this module dependency-free (no React / no DOM) so it can be unit-tested
 * in isolation and the caret math lives in exactly one place.
 */

import {
  DOCUMENT_SYMBOLS,
  DOCUMENT_SYMBOL_GROUPS,
  SYMBOL_PICKER_GROUPS,
  parseRecentSymbols,
  pushRecentSymbol,
  symbolGlyphById,
  type DocumentSymbol,
  type DocumentSymbolCategory,
  type DocumentSymbolGroup,
  type SymbolPickerCategory,
} from "@mrnote/docs/web/components/console/editor/native-document/insertSymbols.ts";

// Re-export the shared catalog under presentation-flavoured names so the
// presentation toolbar imports from a single, slide-local module (and the
// document/presentation pickers can never drift apart — they ARE the same data).
export {
  DOCUMENT_SYMBOLS as PRESENTATION_SYMBOLS,
  DOCUMENT_SYMBOL_GROUPS as PRESENTATION_SYMBOL_GROUPS,
  SYMBOL_PICKER_GROUPS as PRESENTATION_SYMBOL_PICKER_GROUPS,
  parseRecentSymbols as parseRecentPresentationSymbols,
  pushRecentSymbol as pushRecentPresentationSymbol,
  symbolGlyphById,
};
export type {
  DocumentSymbol as PresentationSymbol,
  DocumentSymbolCategory as PresentationSymbolCategory,
  DocumentSymbolGroup as PresentationSymbolGroup,
  SymbolPickerCategory as PresentationSymbolPickerCategory,
};

/**
 * Symbol recents are a user habit rather than deck state. Reuse the document
 * editor's established local preference key so Word and PowerPoint expose the
 * same bounded, whitelist-validated recent row without persisting UI history
 * into either native package.
 */
export const PRESENTATION_RECENT_SYMBOLS_STORAGE_KEY =
  "mrnote:document-symbol-recent";

/** The result of a caret splice: the new text + where the caret should land. */
export interface SymbolInsertResult {
  /** The text with the glyph spliced in. */
  text: string;
  /** The caret offset just AFTER the inserted glyph (UTF-16 code units). */
  caret: number;
}

/**
 * Splice `glyph` into `text`, replacing the `[start, end)` selection (a collapsed
 * caret when `start === end`). Returns the new text and the caret position just
 * after the inserted glyph so the caller can restore the selection.
 *
 * PURE + defensive: `start` / `end` are clamped to `[0, text.length]` and
 * ordered (a backwards selection is normalized), and a non-finite / negative
 * offset falls back to the end of the string (append) so a stale or missing
 * selection can never throw or drop characters. An empty `glyph` is a no-op that
 * still returns a clamped caret. The glyph is inserted as ordinary text — it
 * round-trips through serialize/export like any typed character (no run/style is
 * attached, mirroring the document picker's plain-text insert).
 */
export function insertSymbolIntoText(
  text: string,
  glyph: string,
  start: number,
  end: number,
): SymbolInsertResult {
  const len = text.length;
  const clamp = (value: number): number => {
    if (!Number.isFinite(value)) return len;
    const rounded = Math.floor(value);
    if (rounded < 0) return len;
    return Math.min(rounded, len);
  };
  const a = clamp(start);
  const b = clamp(end);
  const from = Math.min(a, b);
  const to = Math.max(a, b);
  const next = text.slice(0, from) + glyph + text.slice(to);
  return { text: next, caret: from + glyph.length };
}
