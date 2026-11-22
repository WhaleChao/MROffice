/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The decimal and thousands separators of a viewer's region. MRAI's region is
 * the interface language (there is no separate system setting), so this is
 * what Excel's "Use system separators" reads on the viewer's machine: a German
 * viewer draws 5000 through `#,##0.00` as 5.000,00 and an English one as
 * 5,000.00, while the code, the stored value and the file stay the same.
 */
import type { NumberFormatSeparators } from "./number-format";

const CACHE = new Map<string, NumberFormatSeparators | null>();

/**
 * The region's separators as Intl writes them (the group may be a no-break
 * space, U+00A0 or U+202F), or null when Intl cannot read the locale or the
 * region writes no group separator.
 */
export function localeNumberSeparators(locale: string): NumberFormatSeparators | null {
  const cached = CACHE.get(locale);
  if (cached !== undefined) return cached;
  let separators: NumberFormatSeparators | null = null;
  try {
    const parts = new Intl.NumberFormat(locale).formatToParts(12345.6);
    const decimal = parts.find((part) => part.type === "decimal")?.value;
    const group = parts.find((part) => part.type === "group")?.value;
    if (decimal && group && decimal !== group) separators = { decimal, group };
  } catch {
    separators = null;
  }
  CACHE.set(locale, separators);
  return separators;
}

/** A region the page that opened the editor named for its numbers (an
 * ONLYOFFICE integration's editorConfig.region): it wins over the viewer's
 * language, as Excel's region wins over its interface language. */
let regionOverride: string | null = null;

export function setViewerNumberRegion(region: string | null): void {
  regionOverride = region || null;
}

/**
 * The separators a viewer in `locale` draws numbers with, or undefined when
 * they are the code's own "." and "," (en, zh, ja, ko, th, en-GB): the
 * drawing then needs no change at all.
 */
export function viewerNumberSeparators(locale: string | undefined): NumberFormatSeparators | undefined {
  if (regionOverride) locale = regionOverride;
  if (!locale) return undefined;
  const separators = localeNumberSeparators(locale);
  if (!separators || (separators.decimal === "." && separators.group === ",")) return undefined;
  return separators;
}
