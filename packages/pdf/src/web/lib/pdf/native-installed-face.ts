/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A face a write installed on the page (a served substitute's subset, named
 * `TAGTAG+Arimo-Bold` by pdfEmbeddedFontBaseName) read back as the family a
 * reader picks it by: typing into text set in it is set again in that
 * family, its subset installed anew with the characters typed (WPS: text
 * set in a font stays in that font as it is edited).
 */
import { officeFontFamilies, officeFontMetricMatchesWidths, officeFontSubstitution } from "@mrnote/core/web/lib/office-fonts/substitutes.ts";
import { pdfNativeEmbeddedFontKey } from "./native-embedded-font-key";

const STYLE_WEIGHTS: Readonly<Record<string, number>> = Object.freeze({
  Thin: 100, ExtraLight: 200, Light: 300, SemiLight: 350, Regular: 400,
  Medium: 500, SemiBold: 600, Bold: 700, ExtraBold: 800, Black: 900,
});

let bySubstitute: Map<string, string> | null = null;

/** The Office family whose served substitute is `substitute` (stripped of
 * everything but letters and digits, as the installed name holds it): one
 * whose metrics match exactly first. */
function officeFamilyOfSubstitute(substitute: string): string | null {
  if (!bySubstitute) {
    bySubstitute = new Map();
    const exact = new Map<string, string>();
    for (const family of officeFontFamilies()) {
      const substitution = officeFontSubstitution(family);
      const first = substitution?.substitutes[0];
      if (!first) continue;
      const key = first.replace(/[^A-Za-z0-9]/gu, "");
      if (officeFontMetricMatchesWidths(substitution.metric) && !exact.has(key)) exact.set(key, family);
      if (!bySubstitute.has(key)) bySubstitute.set(key, family);
    }
    for (const [key, family] of exact) bySubstitute.set(key, family);
  }
  return bySubstitute.get(substitute) ?? null;
}

/** The embedded-font key that installs the font's family and style again,
 * or null when it is not a face MRAI installed from a served substitute
 * (its resource is named MRAI_EF_..., its base font `[TAG+]Arimo-Bold`). */
export function pdfNativeEmbeddedFontKeyForInstalledFace(
  font: Readonly<{ resourceName: string; baseFont: string }> | null | undefined,
): string | null {
  if (!font?.resourceName.startsWith("MRAI_EF_")) return null;
  const match = /^(?:[A-Z]{6}\+)?([A-Za-z0-9]+)-([A-Za-z]+)$/u.exec(font.baseFont);
  if (!match) return null;
  const family = officeFamilyOfSubstitute(match[1]);
  if (!family) return null;
  const italic = /Italic$/u.test(match[2]);
  const style = match[2].replace(/Italic$/u, "") || "Regular";
  const weight = STYLE_WEIGHTS[style];
  if (weight === undefined) return null;
  return pdfNativeEmbeddedFontKey({ family, bold: weight >= 600, italic });
}
