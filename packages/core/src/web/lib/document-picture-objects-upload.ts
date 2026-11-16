/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { Extension, type AnyExtension } from "@tiptap/core";
import type { Node as ProseMirrorNode } from "@tiptap/pm/model";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import type { EditorView } from "@tiptap/pm/view";

import { uploadPagePicture } from "./page-picture-upload";
import { registerDocumentPictureReferences } from "./document-picture-references";

/** A picture smaller than this stays in the page (an icon costs nothing). */
export const DOCUMENT_PICTURE_OBJECT_MIN_CHARS = 16 * 1024;
const CONCURRENT_UPLOADS = 2;
const RETRY_AFTER_MS = 60_000;

/** Upload one `data:` picture; its `mrai-picture:` token, or null. */
export const uploadDocumentPicture = uploadPagePicture;

/** The `data:` pictures of `doc` worth moving out, by their `src`. */
export function documentDataPictures(doc: ProseMirrorNode): Set<string> {
  const found = new Set<string>();
  doc.descendants((node) => {
    const src = (node.attrs as { src?: unknown } | null)?.src;
    if (typeof src === "string" && src.length >= DOCUMENT_PICTURE_OBJECT_MIN_CHARS && src.startsWith("data:image/")) found.add(src);
    return true;
  });
  return found;
}

/** `view`'s nodes whose `src` is a key of `tokens`, renamed to its token. */
function swapPictures(view: EditorView, tokens: ReadonlyMap<string, string>): boolean {
  const tr = view.state.tr;
  view.state.doc.descendants((node, pos) => {
    const src = (node.attrs as { src?: unknown } | null)?.src;
    const token = typeof src === "string" ? tokens.get(src) : undefined;
    if (token) tr.setNodeMarkup(pos, undefined, { ...node.attrs, src: token });
    return true;
  });
  if (!tr.docChanged) return false;
  // Housekeeping, not an edit: undo never brings the bytes back, and the
  // page ledger does not read it as typing at the seams around the picture
  // (corpus R279-D20: the swap of Figure 3's src, 228 positions from a page
  // seam, marked that seam edited; the paragraph was re-cut a row short and
  // kept two rows on its page where Word keeps three).
  view.dispatch(tr.setMeta("addToHistory", false));
  return true;
}

export function createDocumentPictureObjectsExtension(options: {
  getPageId: () => string | undefined;
  upload?: (pageId: string, dataUrl: string) => Promise<string | null>;
  delayMs?: number;
}): AnyExtension {
  const upload = options.upload ?? uploadDocumentPicture;
  return Extension.create({
    name: "documentPictureObjects",
    addProseMirrorPlugins() {
      return [
        new Plugin({
          key: new PluginKey("documentPictureObjects"),
          view: (view) => {
            const tokens = new Map<string, string>();
            const inFlight = new Set<string>();
            const retryAt = new Map<string, number>();
            let timer: ReturnType<typeof setTimeout> | null = null;
            let destroyed = false;
            const schedule = (delay = options.delayMs ?? 800) => {
              if (timer === null && !destroyed) timer = setTimeout(run, delay);
            };
            const retryLater = () => {
              setTimeout(() => { if (!destroyed) schedule(0); }, RETRY_AFTER_MS + 50);
            };
            const run = () => {
              timer = null;
              const pageId = options.getPageId();
              if (destroyed || !pageId || !view.editable) return;
              const wanted = documentDataPictures(view.state.doc);
              if (wanted.size === 0) return;
              // Pictures already stored: renamed now (again, if a refused
              // transaction or an Undo brought the `data:` URI back).
              if ([...wanted].some((src) => tokens.has(src))) swapPictures(view, tokens);
              for (const src of wanted) {
                if (inFlight.size >= CONCURRENT_UPLOADS) break;
                if (tokens.has(src) || inFlight.has(src) || (retryAt.get(src) ?? 0) > Date.now()) continue;
                inFlight.add(src);
                void upload(pageId, src)
                  .then((token) => {
                    if (!token) { retryAt.set(src, Date.now() + RETRY_AFTER_MS); retryLater(); return; }
                    registerDocumentPictureReferences(pageId, { src: token });
                    tokens.set(src, token);
                  })
                  .catch(() => { retryAt.set(src, Date.now() + RETRY_AFTER_MS); retryLater(); })
                  .finally(() => { inFlight.delete(src); schedule(0); });
              }
            };
            schedule(options.delayMs ?? 1_500);
            return {
              update: (current, previous) => { if (current.state.doc !== previous.doc) schedule(); },
              destroy: () => { destroyed = true; if (timer !== null) clearTimeout(timer); },
            };
          },
        }),
      ];
    },
  });
}
