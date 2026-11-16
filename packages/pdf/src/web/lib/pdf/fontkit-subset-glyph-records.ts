/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { CustomFontSubsetEmbedder } from "pdf-lib";

interface FontkitTrueTypeSubset {
  _addGlyph(gid: number): number;
  glyf: Uint8Array[];
  offset: number;
}

/**
 * fontkit copies raw glyph records into its TrueType subset and then picks the
 * short `loca` format whenever the table is small, but short offsets are
 * stored halved, so any odd-length record (Noto Sans SC has many) shifts every
 * later glyph by one byte and corrupts the subset. Pad each record to a
 * 4-byte boundary before the offsets are computed; zero padding is spec-legal
 * and keeps the embedded program deterministic.
 */
export function alignSubsetGlyphRecords(embedder: CustomFontSubsetEmbedder): void {
  const subset = (embedder as unknown as { subset?: FontkitTrueTypeSubset })
    .subset;
  if (!subset || typeof subset._addGlyph !== "function") return;
  const addGlyph = subset._addGlyph;
  subset._addGlyph = function alignedAddGlyph(
    this: FontkitTrueTypeSubset,
    gid: number,
  ): number {
    const index = addGlyph.call(this, gid);
    const record = this.glyf[index];
    const padding = (4 - (record.length % 4)) % 4;
    if (padding === 0) return index;
    const padded = new Uint8Array(record.length + padding);
    padded.set(record);
    const constructor = record.constructor as {
      from?: (source: Uint8Array) => Uint8Array;
    };
    this.glyf[index] =
      typeof constructor.from === "function" ? constructor.from(padded) : padded;
    this.offset += padding;
    return index;
  };
}
