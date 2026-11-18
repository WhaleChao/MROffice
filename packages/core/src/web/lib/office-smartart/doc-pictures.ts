/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { SmartArtSchemeColor } from "./colors";

/** One diagram picture: its relationship in the data part and its point's own line. */
export interface NativeSmartArtDocPicture {
  readonly imageRelId: string;
  /** The point's own line: `#rrggbb`, or `"none"` for No Outline. */
  readonly outline?: string;
  readonly outlineScheme?: SmartArtSchemeColor;
  readonly outlineWidthPt?: number;
}

export type NativeSmartArtDocPictures = Readonly<Record<string, NativeSmartArtDocPicture>>;

/** A layout node name the page may key a diagram picture by. */
export const SMARTART_DOC_SLOT_NAME = /^[A-Za-z][A-Za-z0-9_]{0,63}$/;

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

function schemeColour(value: unknown): SmartArtSchemeColor | undefined {
  if (!isRecord(value)) return undefined;
  const { scheme, mods } = value;
  if (
    typeof scheme !== "string" ||
    !/^(dk1|lt1|dk2|lt2|tx1|bg1|tx2|bg2|accent[1-6])$/.test(scheme) ||
    !Array.isArray(mods) ||
    mods.length > 12 ||
    !mods.every(
      (mod) =>
        Array.isArray(mod) &&
        mod.length === 2 &&
        typeof mod[0] === "string" &&
        /^[a-zA-Z]{1,24}$/.test(mod[0]) &&
        Number.isInteger(mod[1]) &&
        Math.abs(mod[1]) <= 999_999_999,
    )
  ) {
    return undefined;
  }
  return { scheme, mods: mods.map((mod) => [mod[0], mod[1]] as const) };
}

/**
 * A view's `docPictures`, checked: every entry well formed, or `null` (a page
 * read before row 1005, or one entry malformed). `{}` is a page read under
 * row 1005 whose diagram keeps no picture of its own.
 */
export function nativeSmartArtDocPictures(value: unknown): NativeSmartArtDocPictures | null {
  if (!isRecord(value)) return null;
  const entries = Object.entries(value);
  // Empty: read under row 1005, the diagram holding no picture of its own --
  // which still says how its items' own pictures are drawn.
  if (entries.length > 8) return null;
  const out: Record<string, NativeSmartArtDocPicture> = {};
  for (const [name, raw] of entries) {
    if (!SMARTART_DOC_SLOT_NAME.test(name) || !isRecord(raw)) return null;
    const relId = raw.imageRelId;
    if (typeof relId !== "string" || !relId || relId.length > 128) return null;
    const outline = raw.outline;
    const width = raw.outlineWidthPt;
    const scheme = raw.outlineScheme === undefined ? undefined : schemeColour(raw.outlineScheme);
    if (raw.outlineScheme !== undefined && !scheme) return null;
    if (outline !== undefined && (typeof outline !== "string" || !/^(#[0-9a-f]{6}|none)$/i.test(outline))) return null;
    if (width !== undefined && (typeof width !== "number" || !(width >= 0) || width > 1584)) return null;
    out[name] = {
      imageRelId: relId,
      ...(typeof outline === "string" ? { outline: outline.toLowerCase() } : {}),
      ...(scheme ? { outlineScheme: scheme } : {}),
      ...(typeof width === "number" ? { outlineWidthPt: width } : {}),
    };
  }
  return out;
}
