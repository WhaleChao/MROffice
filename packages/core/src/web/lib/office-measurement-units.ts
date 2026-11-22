/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Office measurement system for an app locale.
 *
 * Word, Excel and PowerPoint show lengths in the operating system's
 * measurement system (Microsoft Support, "Change the measurement units from
 * inches to centimeters": Windows Region > Measurement system, macOS Language
 * & Region > Measurement system). A web page cannot read that setting, so the
 * app locale stands in for the region: US English (this app's "en"; "en-US")
 * is the US system, inches; every other locale, en-GB included, is metric.
 *
 * Moved verbatim from the Spreadsheet Page Setup dialog so the Presentation
 * editor shares the same rule rather than a third one. (Docs'
 * `documentTableMeasurementUnit` matches only a literal "en-US" and so reads
 * the app's "en" as metric; it is left to its owner.)
 */
/**
 * A measurement system the host fixed for this editor (an ONLYOFFICE
 * config's customization.unit, lib/embed/onlyoffice-frame): it wins over the
 * locale and the browser's region, as ONLYOFFICE's unit wins over its
 * region. null: none fixed.
 */
let measurementOverride: "cm" | "inch" | null = null;

export function setOfficeMeasurementOverride(unit: "cm" | "inch" | null): void {
  measurementOverride = unit;
}

export function officeLocaleUsesInches(locale: string): boolean {
  if (measurementOverride) return measurementOverride === "inch";
  try {
    const tag = new Intl.Locale(locale);
    return tag.language === "en" && (tag.region ?? "US") === "US";
  } catch {
    return false;
  }
}

/** Regions whose system measurement is US customary (CLDR): the US, Liberia, Myanmar. */
const OFFICE_INCH_REGIONS = new Set(["US", "LR", "MM"]);

/**
 * Whether the SYSTEM measures in inches, as far as a browser can tell.
 * PowerPoint's rulers follow the operating system's region setting, which a
 * page cannot read; the region of the browser's own language
 * (`navigator.language`, e.g. "en-GB", "zh-CN") is the nearest it states:
 * inches in the US, Liberia and Myanmar, centimetres everywhere else. A
 * language with no region falls back to the interface locale's rule.
 */
export function officeSystemUsesInches(
  appLocale: string,
  browserLanguage: string | undefined = typeof navigator === "undefined" ? undefined : navigator.language,
): boolean {
  if (measurementOverride) return measurementOverride === "inch";
  if (browserLanguage) {
    try {
      const region = new Intl.Locale(browserLanguage).maximize().region;
      if (region) return OFFICE_INCH_REGIONS.has(region);
    } catch {
      // an unreadable tag: the interface's rule below
    }
  }
  return officeLocaleUsesInches(appLocale);
}
