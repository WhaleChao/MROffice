/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Insert Symbol (Excel / WPS "Insert ▸ Symbol") — a curated character picker for
 * cells. Pure data + insertion math so the picker grid and the
 * insert-at-caret logic are unit-testable without a grid.
 *
 * Excel's Symbol dialog is a large font-driven grid; this is a practical curated
 * subset grouped by category (currency, math, arrows, punctuation, Greek,
 * symbols/marks). Each entry is a single grapheme the user clicks to insert into
 * the active cell at the caret. The set is intentionally small and static (no
 * font enumeration) so it renders instantly and stays deterministic.
 */

/** One symbol category shown as a section in the picker. */
export interface SymbolCategory {
  /** Stable id (used as the section key / i18n suffix). */
  readonly id: string;
  /** The characters in this category, in display order. */
  readonly symbols: readonly string[];
}

/**
 * The curated symbol categories. Kept compact (a few rows each) so the whole
 * panel fits without scrolling on a normal window. Order is currency → math →
 * arrows → punctuation → Greek → marks (most-used first).
 */
export const SYMBOL_CATEGORIES: readonly SymbolCategory[] = [
  {
    id: "currency",
    symbols: ["€", "£", "¥", "¢", "₩", "₹", "₽", "₺", "₫", "฿", "₴", "₦"],
  },
  {
    id: "math",
    symbols: [
      "±", "×", "÷", "≈", "≠", "≤", "≥", "∞", "√", "∑", "∏", "∫",
      "µ", "π", "Δ", "∂", "°", "‰", "½", "¼", "¾", "⅓", "⅔", "²",
      "³", "·", "∝", "∈", "∉", "∅", "∩", "∪",
    ],
  },
  {
    id: "arrows",
    symbols: ["←", "→", "↑", "↓", "↔", "↕", "⇐", "⇒", "⇔", "↗", "↘", "↩"],
  },
  {
    id: "punctuation",
    symbols: [
      "“", "”", "‘", "’", "«", "»", "—", "–", "…", "•", "·", "§",
      "¶", "†", "‡", "·", "¡", "¿",
    ],
  },
  {
    id: "greek",
    symbols: [
      "α", "β", "γ", "δ", "ε", "θ", "λ", "μ", "ρ", "σ", "φ", "ω",
      "Γ", "Δ", "Θ", "Λ", "Σ", "Φ", "Ω", "Π",
    ],
  },
  {
    id: "marks",
    symbols: [
      "©", "®", "™", "✓", "✗", "★", "☆", "☑", "☒", "→", "♥", "♦",
      "♣", "♠", "№", "℃", "℉", "⚠",
    ],
  },
] as const;

/** A flat, de-duplicated list of every symbol across all categories. */
export function allSymbols(): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const category of SYMBOL_CATEGORIES) {
    for (const symbol of category.symbols) {
      if (seen.has(symbol)) continue;
      seen.add(symbol);
      out.push(symbol);
    }
  }
  return out;
}

/**
 * Insert `symbol` into `value` at the caret, replacing any selected range
 * [selStart, selEnd). Returns the new value AND the caret offset just after the
 * inserted symbol so the editor can restore the caret. Selection bounds are
 * clamped into [0, value.length] and ordered, so a stale/null-ish selection
 * still inserts at a sane position (the symbol lands at the clamped caret).
 *
 * The symbol is inserted verbatim; a multi-code-unit grapheme (e.g. an emoji)
 * advances the caret by its full `.length` so the caret never lands mid-surrogate.
 */
export function insertSymbolAt(
  value: string,
  symbol: string,
  selStart: number,
  selEnd: number,
): { value: string; caret: number } {
  const len = value.length;
  const clamp = (n: number): number => {
    if (!Number.isFinite(n)) return len;
    return Math.max(0, Math.min(len, Math.floor(n)));
  };
  const a = clamp(selStart);
  const b = clamp(selEnd);
  const start = Math.min(a, b);
  const end = Math.max(a, b);
  const next = value.slice(0, start) + symbol + value.slice(end);
  return { value: next, caret: start + symbol.length };
}
