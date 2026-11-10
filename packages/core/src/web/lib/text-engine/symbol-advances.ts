/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import table from "../office-fonts/symbol-advances/symbol-advances.json";

interface SymbolFace {
  readonly unitsPerEm: number;
  readonly first: number;
  readonly advances: readonly (number | null)[];
}

const faces = new Map<string, SymbolFace>(
  Object.entries((table as { faces: Record<string, SymbolFace> }).faces).map(([family, face]) => [family.toLowerCase(), face]),
);

/** True when `family` names one of the recorded symbol faces. */
export function isRecordedSymbolFace(family: string): boolean {
  return faces.has(family.trim().replace(/^["']|["']$/g, "").toLowerCase());
}

/**
 * The advance, in em, of `char` in the symbol face `family`; undefined when the
 * face is not recorded or has no glyph there. A symbol font's character is a
 * byte, stored as itself (U+0020-U+00FF) or shifted into the private use area
 * (U+F020-U+F0FF), as Office writes it.
 */
export function symbolAdvanceEm(family: string, char: string): number | undefined {
  const face = faces.get(family.trim().replace(/^["']|["']$/g, "").toLowerCase());
  if (!face || char.length !== 1) return undefined;
  let code = char.charCodeAt(0);
  if (code >= 0xf020 && code <= 0xf0ff) code -= 0xf000;
  const advance = face.advances[code - face.first];
  return advance === undefined || advance === null ? undefined : advance / face.unitsPerEm;
}
