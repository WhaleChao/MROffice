/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { fetchVerifiedPdfUnicodeFallbackFont } from "./unicode-font-asset";

/**
 * The CSS family of the East Asian face a saved text box is drawn in: the
 * same Noto Sans SC file the server embeds (pdf_fonts.py, pinned by size and
 * SHA-256). The app does not ship that face as a web font, so naming
 * "Noto Sans SC" alone fell back to the system's CJK face: the draft and the
 * read view drew other glyphs, with other widths, than the page the save
 * paints. Registered from the verified bytes, the editor, the read view and
 * the row measurement set the same glyphs the PDF does.
 */
export const PDF_EAST_ASIAN_CSS_FAMILY = "MRNotePdfCJK";

type FaceStatus = "idle" | "loading" | "available" | "unavailable";
type FaceState = { status: FaceStatus; promise?: Promise<boolean>; listeners: Set<() => void> };
const documents = new WeakMap<Document, FaceState>();

function stateFor(doc: Document): FaceState {
  let state = documents.get(doc);
  if (!state) {
    state = { status: "idle", listeners: new Set() };
    documents.set(doc, state);
  }
  return state;
}

export function pdfEastAsianFaceStatus(doc?: Document | null): FaceStatus {
  const target = doc ?? (typeof document === "undefined" ? null : document);
  return target ? stateFor(target).status : "unavailable";
}

/** Told when the face becomes available (or is found unavailable). */
export function subscribePdfEastAsianFace(doc: Document, listener: () => void): () => void {
  const state = stateFor(doc);
  state.listeners.add(listener);
  return () => { state.listeners.delete(listener); };
}

/**
 * Register the face once per document. Resolves false, without throwing,
 * where it cannot be loaded (no FontFace API, the asset missing or not the
 * pinned bytes): the stack then falls back as it did before.
 */
export function ensurePdfEastAsianFace(doc?: Document | null): Promise<boolean> {
  const target = doc ?? (typeof document === "undefined" ? null : document);
  if (!target || typeof FontFace === "undefined" || !target.fonts) return Promise.resolve(false);
  const state = stateFor(target);
  if (state.promise) return state.promise;
  state.status = "loading";
  state.promise = (async () => {
    try {
      const bytes = await fetchVerifiedPdfUnicodeFallbackFont();
      if (!bytes) throw new Error("pdf_east_asian_face_unavailable");
      // Declared as the one regular upright face it is: a bold or an italic
      // asked of it is synthesised by the browser (a heavier stroke, a 12°
      // lean), as WPS and Word synthesise them for a face without its own.
      // The text box's appearance is the browser's own raster of the draft,
      // and a page run set in this face draws its bold stroked and its
      // italic leaning the same way (paragraph-rewrite.ts), so the draft,
      
      // was declared for every weight and style so nothing was synthesised,
      // and a word made bold and italic read so on the ribbon but drew
      // regular; owner's stress page, 22:21.)
      const face = new FontFace(PDF_EAST_ASIAN_CSS_FAMILY, bytes, {
        style: "normal", weight: "400", display: "block",
      });
      await face.load();
      target.fonts.add(face);
      state.status = "available";
      return true;
    } catch {
      state.status = "unavailable";
      return false;
    } finally {
      for (const listener of state.listeners) {
        try { listener(); } catch { /* a listener never breaks the load */ }
      }
    }
  })();
  return state.promise;
}
