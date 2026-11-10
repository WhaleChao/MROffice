/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A font-family list as CSS will read it.
 *
 * CSS takes a family name unquoted only when every word of it is an
 * identifier: a word may not start with a digit (or a hyphen and a digit) nor
 * hold punctuation. "Wingdings 3", "Bodoni 72" or a theme reference such as
 * "+mn-lt" written bare make the WHOLE declaration invalid -- a style object,
 * an inline style, an SVG `font-family` attribute and a canvas `font` string
 * all drop it -- and the text draws in the inherited face (a Facet bullet as
 * a notdef box). Names from documents therefore go through here on their way
 * into CSS: every family that cannot stand bare is quoted, generic families
 * and CSS-wide keywords stay keywords, and a name already quoted is kept.
 */

const KEYWORDS = new Set([
  "serif",
  "sans-serif",
  "monospace",
  "cursive",
  "fantasy",
  "system-ui",
  "ui-serif",
  "ui-sans-serif",
  "ui-monospace",
  "ui-rounded",
  "emoji",
  "math",
  "fangsong",
  "inherit",
  "initial",
  "unset",
  "revert",
  "revert-layer",
  "-apple-system",
  "blinkmacsystemfont",
]);

const IDENTIFIER = /^(?:--|-?[A-Za-z_\u0080-￿])[-A-Za-z0-9_\u0080-￿]*$/;

/** Split a family list on the commas between families (not inside quotes). */
export function splitFontFamilyList(value: string): string[] {
  const out: string[] = [];
  let current = "";
  let quote: string | null = null;
  for (let i = 0; i < value.length; i += 1) {
    const ch = value[i];
    if (quote) {
      current += ch;
      if (ch === "\\" && i + 1 < value.length) {
        current += value[i + 1];
        i += 1;
      } else if (ch === quote) quote = null;
    } else if (ch === '"' || ch === "'") {
      quote = ch;
      current += ch;
    } else if (ch === ",") {
      out.push(current.trim());
      current = "";
    } else current += ch;
  }
  out.push(current.trim());
  return out.filter((item) => item.length > 0);
}

function quoted(item: string): boolean {
  return item.length >= 2 && (item[0] === '"' || item[0] === "'") && item[item.length - 1] === item[0];
}

/** Whether one family stands in CSS as written. */
export function cssFontFamilyItemIsValid(item: string): boolean {
  const trimmed = item.trim();
  if (!trimmed) return false;
  if (quoted(trimmed)) return true;
  if (KEYWORDS.has(trimmed.toLowerCase())) return true;
  // A CSS function (`var(--font-sans)`) is the stylesheet's own value.
  if (/^[-A-Za-z]+\(.*\)$/.test(trimmed)) return true;
  return trimmed.split(/\s+/).every((word) => IDENTIFIER.test(word));
}

/** Whether a whole family list stands in CSS as written. */
export function cssFontFamilyIsValid(value: string): boolean {
  const items = splitFontFamilyList(value);
  return items.length > 0 && items.every(cssFontFamilyItemIsValid);
}

/** One family as CSS reads it: quoted unless it can stand bare. */
export function cssFontFamilyName(name: string): string {
  const trimmed = name.trim();
  if (cssFontFamilyItemIsValid(trimmed)) return trimmed;
  return `"${trimmed.replace(/\\/g, "\\\\").replace(/"/g, '\\"')}"`;
}

/** A family list as CSS reads it, each family that cannot stand bare quoted. */
export function cssFontFamilyList(value: string): string {
  return splitFontFamilyList(value).map(cssFontFamilyName).join(", ");
}

/** A family list as CSS reads it, or the empty / absent value it was. */
export function cssFontFamilyOrUndefined<T extends string | null | undefined>(value: T): T | string {
  return typeof value === "string" && value.trim() ? cssFontFamilyList(value) : value;
}
