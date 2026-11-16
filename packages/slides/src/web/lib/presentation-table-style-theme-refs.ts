/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { normalizeHexColor } from "@mrnote/core/web/lib/office-color-palette.ts";
import {
  normalizeOfficeThemeColorRef,
  officeThemeRefHex,
  type OfficeThemeGridColors,
} from "@mrnote/core/web/lib/office-theme-color-grid.ts";

type Rec = Record<string, unknown>;
const record = (value: unknown): Rec | null =>
  value !== null && typeof value === "object" && !Array.isArray(value) ? (value as Rec) : null;

/** Where a reference's colour lives in its part: `color`, `fill`, `borders.<edge>.color`. */
function colourAt(part: Rec, path: string): string | null {
  const steps = path.split(".");
  let node: unknown = part;
  for (const step of steps) node = record(node)?.[step];
  return typeof node === "string" ? normalizeHexColor(node) : null;
}

function withColourAt(part: Rec, path: string, colour: string): Rec {
  const [head, ...rest] = path.split(".");
  if (!rest.length) return { ...part, [head]: colour };
  const child = record(part[head]);
  return child ? { ...part, [head]: withColourAt(child, rest.join("."), colour) } : part;
}

/** Whether a table style look carries any theme reference. */
export function tableStyleLookHasThemeRefs(look: unknown): boolean {
  const parts = record(record(look)?.parts);
  return !!parts && Object.values(parts).some((part) => record(record(part)?.themeColors) !== null);
}

/**
 * `look` with every colour its theme references name drawn in `colors` (the
 * deck's Theme Colors grid), each reference's `rgb` the new colour; the same
 * look when nothing moves.
 */
export function tableStyleLookUnderTheme<L>(look: L, colors: OfficeThemeGridColors): L {
  const value = record(look);
  const parts = record(value?.parts);
  if (!value || !parts) return look;
  let changed = false;
  const nextParts: Rec = {};
  for (const [name, raw] of Object.entries(parts)) {
    let part = record(raw);
    const refs = record(part?.themeColors);
    if (!part || !refs) {
      nextParts[name] = raw;
      continue;
    }
    const nextRefs: Rec = { ...refs };
    let partChanged = false;
    for (const [path, rawRef] of Object.entries(refs)) {
      const ref = normalizeOfficeThemeColorRef(rawRef);
      const stated = normalizeHexColor(record(rawRef)?.rgb);
      const current = colourAt(part, path);
      if (!ref || !stated || current !== stated) continue;
      const next = officeThemeRefHex(ref, colors);
      if (!next || next === current) continue;
      part = withColourAt(part, path, next);
      nextRefs[path] = { ...(rawRef as Rec), rgb: next };
      partChanged = true;
    }
    if (partChanged) changed = true;
    nextParts[name] = partChanged ? { ...part, themeColors: nextRefs } : raw;
  }
  return changed ? ({ ...value, parts: nextParts } as L) : look;
}
