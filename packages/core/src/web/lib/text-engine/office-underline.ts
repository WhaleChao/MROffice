/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import table from "../office-fonts/office-underline.json";

type Metrics = readonly [position: number, thickness: number, unitsPerEm: number];
const families = (table as unknown as { families: Record<string, Record<string, Metrics>> }).families;

export interface OfficeUnderline {
  /** From the baseline down to the line's top, pt. */
  readonly topPt: number;
  readonly thicknessPt: number;
}

/** The underline of a run in `stack` (its first recorded family) at `sizePt`, or null when no family is recorded. */
export function officeUnderline(stack: string, weight: number, italic: boolean, sizePt: number): OfficeUnderline | null {
  for (const raw of stack.split(",")) {
    const styles = families[raw.trim().replace(/^["']|["']$/g, "").toLowerCase()];
    if (!styles) continue;
    const bold = weight >= 600 ? "700" : "400";
    const metrics = styles[`${bold}${italic ? "i" : ""}`] ?? styles[bold] ?? styles["400"] ?? Object.values(styles)[0];
    if (!metrics) return null;
    const [position, thickness, unitsPerEm] = metrics;
    return {
      topPt: Math.round((-position / unitsPerEm) * sizePt),
      thicknessPt: Math.max(1, Math.round((thickness / unitsPerEm) * sizePt)),
    };
  }
  return null;
}
