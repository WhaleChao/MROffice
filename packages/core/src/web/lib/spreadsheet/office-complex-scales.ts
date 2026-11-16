/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// The widths Windows Excel sets Arabic, Hebrew and Bengali letters at, by the
// cell's face (office-complex-scales.json: Excel's printed width over the
// packaged Noto face of the script, lookwin p84-lw and p90-lw). The API's
// sheet PDF writer prints the letters a scale lists in that Noto face at the
// scaled size (notebooks.py complex_pieces, office_complex_scales.json, the
// same table); a cell's stack here carries the same face at the same scale
// right behind the face's own stand-ins, so the grid, its row fitting and the
// editor's print set them as wide as Excel does rather than in whatever face
// further down the stack has them (Wine Tahoma's Arabic, behind every face,
// or the script fallback's Noto at its own size: 20-30% wide in Calibri,
// Cambria, Aptos and Verdana cells, 8% in Bengali).
import table from "./office-complex-scales.json";
import { ensureOfficeScaledScriptFamily } from "../office-fonts/substitutes";

export type OfficeComplexScript = "arabic" | "hebrew" | "bengali";

interface ScriptFace {
  readonly script: OfficeComplexScript;
  /** The packaged face the scale is measured against. */
  readonly face: string;
  /** A letter of the script: the face's slices holding it are declared. */
  readonly letter: number;
  /** The characters the API's writer reads as the script's (notebooks.py _complex_script_of). */
  readonly range: string;
}

const SCRIPT_FACES: readonly ScriptFace[] = [
  { script: "arabic", face: "Noto Sans Arabic", letter: 0x0627, range: "U+0600-08FF, U+FB50-FDFF, U+FE70-FEFF" },
  { script: "hebrew", face: "Noto Sans Hebrew", letter: 0x05d0, range: "U+0590-05FF, U+FB1D-FB4F" },
  { script: "bengali", face: "Noto Sans Bengali", letter: 0x0995, range: "U+0980-09FF" },
];

/**
 * The scripts a face's own stack already sets at Office's advances: Arial's
 * Arabic and Hebrew are metric clones of Office's own (catalog scriptParts,
 * MRAI-Metric-Arial-Arabic / -Hebrew), declared under the family itself.
 */
const OWN_SCRIPT_PARTS: Readonly<Record<string, readonly OfficeComplexScript[]>> = {
  arial: ["arabic", "hebrew"],
};

const scales = table.scales as Readonly<Record<OfficeComplexScript, Readonly<Record<string, number>>>>;

/** Excel's width of `script`'s letters in a cell of `face` over the packaged Noto face's; undefined where unmeasured. */
export function officeComplexScale(script: OfficeComplexScript, face: string | null | undefined): number | undefined {
  const byFace = scales[script];
  if (!byFace) return undefined;
  const key = (face ?? "").trim().replace(/^["']|["']$/g, "").trim().toLowerCase();
  return byFace[key] ?? byFace["*"];
}

/** What every family {@link officeComplexScaleFamily} names starts with: not an Office family to look up. */
export const OFFICE_COMPLEX_SCALE_PREFIX = "MRAI Excel ";

/** The family setting `script` at `scale`: one per script and scale, shared by every face that scale is measured for. */
export function officeComplexScaleFamily(script: OfficeComplexScript, scale: number): string {
  return `${OFFICE_COMPLEX_SCALE_PREFIX}${script} ${Math.round(scale * 1000)}`;
}

/**
 * The quoted families, comma-terminated, a cell stack in `face` carries
 * behind the face and its stand-ins: each listed script's letters in its
 * packaged Noto face at Excel's scale. Each family is declared on first use.
 */
export function officeComplexScaleStack(face: string | null | undefined): string {
  const key = (face ?? "").trim().toLowerCase();
  let out = "";
  for (const entry of SCRIPT_FACES) {
    if (OWN_SCRIPT_PARTS[key]?.includes(entry.script)) continue;
    const scale = officeComplexScale(entry.script, face);
    if (scale === undefined || !(scale > 0)) continue;
    const family = officeComplexScaleFamily(entry.script, scale);
    ensureOfficeScaledScriptFamily(family, entry.face, entry.letter, entry.range, `${(scale * 100).toFixed(2)}%`);
    out += `"${family}", `;
  }
  return out;
}
