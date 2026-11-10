/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The served font slices a PDF text draft needs, fetched and verified by the
 * editor (lib/office-fonts/embeddable-faces.ts): every slice of every face a
 * family the draft sets text in could draw one of the text's characters
 * from. The engine host plans the write with exactly these (it fetches
 * nothing) and checks each against the lock again.
 */

import {
  EmbeddableFontError,
  embeddableFontFamily,
  embeddableSlicesFor,
  fetchEmbeddableFontSlice,
  loadEmbeddableFontManifest,
  selectEmbeddableFontFace,
  type EmbeddableFontManifest,
  type EmbeddableFontSlice,
} from "@mrnote/core/web/lib/office-fonts/embeddable-faces.ts";
import {
  pdfNativeEmbeddedFontRefusalTag,
  pdfNativeTextDraftEmbeddedFontRequests,
} from "./native-embedded-font-key";
import { choosePdfPageFont, type PdfPageFont } from "./page-fonts";

export interface PdfNativeEmbeddedFontProgram {
  readonly sha256: string;
  readonly bytes: Uint8Array;
  /** A font from the reader's computer the page keeps: the face the
   * program is (native-embedded-fonts.ts installs it as its own slice). */
  readonly pageFont?: PdfPageFont;
}

/** One write never carries more than this much font data to the engine. */
const MAX_PROGRAM_BYTES = 32 * 1024 * 1024;









export async function fetchPdfNativeEmbeddedFontManifest(
  draft: Readonly<{ fontResourceName: string; segments?: readonly Readonly<{ fontResourceName?: string }>[] }> | null | undefined,
): Promise<EmbeddableFontManifest | null> {
  const substitutes = [...new Set(pdfNativeTextDraftEmbeddedFontRequests(draft)
    .flatMap((request) => embeddableFontFamily(request.family)?.substitutes ?? []))];
  if (substitutes.length === 0) return null;
  const manifest = await loadEmbeddableFontManifest(substitutes);
  return Object.freeze({
    version: manifest.version,
    ranges: manifest.ranges,
    families: Object.freeze(Object.fromEntries(substitutes.flatMap((substitute) =>
      manifest.families[substitute] ? [[substitute, manifest.families[substitute]]] : []))),
  });
}

export async function fetchPdfNativeEmbeddedFontPrograms(
  draft: Readonly<{ fontResourceName: string; segments?: readonly Readonly<{ fontResourceName?: string }>[] }> | null | undefined,
  text: string,
  options?: Readonly<{
    fetch?: typeof globalThis.fetch;
    /** The page's fonts from readers' computers, and how their programs
     * are read (the page's font attachments). */
    pageFonts?: readonly PdfPageFont[];
    loadPageFont?: (font: PdfPageFont) => Promise<Uint8Array>;
  }>,
): Promise<readonly PdfNativeEmbeddedFontProgram[]> {
  const requests = pdfNativeTextDraftEmbeddedFontRequests(draft);
  if (requests.length === 0) return [];
  // A family MRAI does not serve, set in the page's own face of it.
  const pagePrograms: PdfNativeEmbeddedFontProgram[] = [];
  for (const request of requests) {
    if (embeddableFontFamily(request.family)) continue;
    const choice = choosePdfPageFont(options?.pageFonts ?? [], request.family, request);
    if (!choice || !options?.loadPageFont || pagePrograms.some((program) => program.sha256 === choice.font.sha256)) continue;
    try {
      pagePrograms.push(Object.freeze({ sha256: choice.font.sha256, bytes: await options.loadPageFont(choice.font), pageFont: choice.font }));
    } catch (error) {
      throw new Error(`${pdfNativeEmbeddedFontRefusalTag({ code: "unavailable", family: request.family })} ${error instanceof Error ? error.message : ""}`.trim());
    }
  }
  // Only the faces of the families this draft names (the manifest is served
  // one family at a time).
  const substitutes = requests.flatMap((request) => embeddableFontFamily(request.family)?.substitutes ?? []);
  if (substitutes.length === 0) return pagePrograms;
  let manifest: Awaited<ReturnType<typeof loadEmbeddableFontManifest>>;
  try {
    manifest = await loadEmbeddableFontManifest(substitutes);
  } catch (error) {
    if (error instanceof EmbeddableFontError) {
      const family = requests[0]?.family ?? "";
      throw new Error(`${pdfNativeEmbeddedFontRefusalTag({ code: "unavailable", family })} ${error.message}`);
    }
    throw error;
  }
  const codePoints = [...new Set([...text].map((character) => character.codePointAt(0)!))];
  const slices = new Map<string, EmbeddableFontSlice>();
  for (const request of requests) {
    const family = embeddableFontFamily(request.family);
    // A family MRAI does not serve is refused by the planner, in words.
    if (!family) continue;
    for (const substitute of family.substitutes) {
      const choice = selectEmbeddableFontFace(manifest, family, substitute, request);
      if (!choice) continue;
      for (const codePoint of codePoints) {
        for (const slice of embeddableSlicesFor(choice.face, codePoint)) slices.set(slice.sha256, slice);
      }
    }
  }
  const total = [...slices.values()].reduce((sum, slice) => sum + slice.size, 0);
  const family = requests[0]?.family ?? "";
  if (total > MAX_PROGRAM_BYTES) {
    throw new Error(`${pdfNativeEmbeddedFontRefusalTag({ code: "unavailable", family })} The text needs too many font files at once.`);
  }
  try {
    return [...pagePrograms, ...await Promise.all([...slices.values()].map(async (slice) => Object.freeze({
      sha256: slice.sha256,
      // A slice is shared by every draft that needs it: one draft's abort
      // never cancels the fetch another is waiting for.
      bytes: await fetchEmbeddableFontSlice(slice, { fetch: options?.fetch }),
    })))];
  } catch (error) {
    if (error instanceof EmbeddableFontError) {
      throw new Error(`${pdfNativeEmbeddedFontRefusalTag({ code: "unavailable", family })} ${error.message}`);
    }
    throw error;
  }
}
