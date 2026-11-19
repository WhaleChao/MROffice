/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import table from "./powerpoint-script-fallbacks.json";
import { ensureOfficeRangedAliasFamily } from "./substitutes";

type FallbackFace = "Calibri" | "Cambria";
const FACES: ReadonlyMap<string, FallbackFace> = new Map(
  Object.entries((table as { faces: Record<string, { face: FallbackFace }> }).faces).map(([family, entry]) => [family.toLowerCase(), entry.face] as const),
);

/** The scripts the probe measured: Latin and its extensions (Vietnamese among them), combining marks, Greek, Cyrillic. */
export const POWERPOINT_SCRIPT_FALLBACK_RANGES: readonly (readonly [number, number])[] = [
  [0x0000, 0x024f],
  [0x0300, 0x036f],
  [0x0370, 0x03ff],
  [0x0400, 0x052f],
  [0x1e00, 0x1eff],
  [0x1f00, 0x1fff],
];

/** Each replacing face's served cuts, Office's own metrics last (tried first where both draw a letter). */
const SERVED: Readonly<Record<FallbackFace, { alias: string; families: readonly string[] }>> = {
  Calibri: { alias: "MRAI PowerPoint Fallback Calibri", families: ["Carlito"] },
  Cambria: { alias: "MRAI PowerPoint Fallback Cambria", families: ["Caladea", "MRAI-Metric-Cambria"] },
};

/** The face PowerPoint for Windows draws `family`'s missing letters in, when measured. */
export function powerPointScriptFallbackFace(family: string): FallbackFace | undefined {
  return FACES.get(family.trim().replace(/^["']|["']$/g, "").trim().toLowerCase());
}

/**
 * `stack` with, behind each measured Office face's metric clone
 * ("MRAI-Metric-…", as `withOfficeFontSubstitutes` lists it), the family that
 * draws the face's missing Latin, Greek and Cyrillic letters as PowerPoint
 * for Windows does. A stack without such a face is returned as it is.
 */
export function withPowerPointScriptFallbacks(stack: string): string {
  const parts = stack.split(",").map((part) => part.trim()).filter(Boolean);
  const unquote = (part: string) => part.replace(/^["']|["']$/g, "").trim();
  const out: string[] = [];
  let pending: FallbackFace | undefined;
  let changed = false;
  for (const part of parts) {
    const name = unquote(part);
    const face = powerPointScriptFallbackFace(name);
    // Cambria's own missing letters: no other face of PowerPoint's to fall to.
    if (face) pending = face.toLowerCase() === name.toLowerCase() ? undefined : face;
    out.push(part);
    if (pending && name.startsWith("MRAI-Metric-") && !face) {
      const served = SERVED[pending];
      ensureOfficeRangedAliasFamily(served.alias, served.families, POWERPOINT_SCRIPT_FALLBACK_RANGES);
      if (!out.some((existing) => unquote(existing) === served.alias)) out.push(`"${served.alias}"`);
      pending = undefined;
      changed = true;
    }
  }
  return changed ? out.join(", ") : stack;
}
