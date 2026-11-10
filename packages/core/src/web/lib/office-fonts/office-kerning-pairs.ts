/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * One recorded style of Office's pair kerning (office-kerning/*.json) as a lookup,
 * in ems. Shared by the text engine (a served stand-in's kerning) and the SmartArt
 * engine's fit (`ms-engine/latin-face`).
 */

export interface OfficeKerningStyle {
  readonly unitsPerEm: number;
  readonly left: readonly string[];
  readonly right: readonly string[];
  readonly pairs: readonly string[];
}

/** Font units by left and right code point. */
export function compileOfficeKerningUnits(style: OfficeKerningStyle): (left: number, right: number) => number {
  const leftClass = new Map<number, number>();
  style.left.forEach((members, n) => members.split(" ").forEach((hex) => leftClass.set(parseInt(hex, 16), n)));
  const rightClass = new Map<number, number>();
  style.right.forEach((members, n) => members.split(" ").forEach((hex) => rightClass.set(parseInt(hex, 16), n)));
  const rows = style.pairs.map((row) => {
    const values = new Map<number, number>();
    for (const entry of row.split(" ")) {
      if (!entry) continue;
      const [cls, value] = entry.split(":");
      values.set(Number(cls), Number(value));
    }
    return values;
  });
  return (left, right) => {
    const row = leftClass.get(left);
    const column = rightClass.get(right);
    return row === undefined || column === undefined ? 0 : rows[row]?.get(column) ?? 0;
  };
}

/** The same in ems. */
export function compileOfficeKerningStyle(style: OfficeKerningStyle): (left: number, right: number) => number {
  const units = compileOfficeKerningUnits(style);
  return (left, right) => units(left, right) / style.unitsPerEm;
}
