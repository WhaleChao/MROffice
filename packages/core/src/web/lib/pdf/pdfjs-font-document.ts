/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
type FontSetLike = Pick<FontFaceSet, "add" | "delete" | "has">;

interface Waiting {
  face: FontFace;
  settle: () => void;
}

const FLUSH_FALLBACK_MS = 50;

class DeferredFontSet implements FontSetLike {
  private readonly waiting = new Set<FontFace>();
  private queue: Waiting[] = [];
  private flushScheduled = false;

  constructor(private readonly fonts: FontFaceSet) {}

  add(face: FontFace): FontFaceSet {
    if (this.waiting.has(face) || this.fonts.has(face)) return this as unknown as FontFaceSet;
    if (face.status === "loaded" || face.status === "error") {
      this.fonts.add(face);
      return this as unknown as FontFaceSet;
    }
    this.waiting.add(face);
    const native = face.loaded;
    const nativeLoad = face.load;
    const joined = new Promise<FontFace>((resolve, reject) => {
      native.then(
        () => this.enqueue({ face, settle: () => resolve(face) }),
        (error: unknown) => this.enqueue({ face, settle: () => reject(error) }),
      );
    });
    // PDF.js awaits the promise; this branch only keeps an unawaited
    // rejection (a face dropped before it loaded) from being reported.
    joined.catch(() => {});
    // What PDF.js waits on: `loaded` (an embedded face) or `load()` (a
    // system face it loads itself). Both now settle once the face joined.
    Object.defineProperty(face, "loaded", { configurable: true, get: () => joined });
    Object.defineProperty(face, "load", {
      configurable: true,
      writable: true,
      value: () => {
        void nativeLoad.call(face).catch(() => {});
        return joined;
      },
    });
    return this as unknown as FontFaceSet;
  }

  delete(face: FontFace): boolean {
    const wasWaiting = this.waiting.delete(face);
    // A face a held drawing still sets text in (an editor's draft, the words
    // a line left a moment ago keeps on the page) leaves once nothing holds
    // it: PDF.js clears a document's faces when the page reads the next
    
    
    const family = fontFamilyKey(String(face.family ?? ""));
    if ((fontHolds.get(family) ?? 0) > 0 && this.fonts.has(face)) {
      heldDeletions.add({ fonts: this.fonts, face, family });
      return true;
    }
    return this.fonts.delete(face) || wasWaiting;
  }

  has(face: FontFace): boolean {
    return this.waiting.has(face) || this.fonts.has(face);
  }

  private enqueue(item: Waiting): void {
    this.queue.push(item);
    if (this.flushScheduled) return;
    this.flushScheduled = true;
    let done = false;
    const flush = () => {
      if (done) return;
      done = true;
      this.flushScheduled = false;
      const batch = this.queue;
      this.queue = [];
      // Every face first, then the waiters: the set changes once.
      for (const { face } of batch) {
        if (this.waiting.delete(face)) this.fonts.add(face);
      }
      for (const { settle } of batch) settle();
    };
    // The next frame, or soon in a tab that draws no frames.
    if (typeof requestAnimationFrame === "function") requestAnimationFrame(flush);
    setTimeout(flush, FLUSH_FALLBACK_MS);
  }
}

const fontHolds = new Map<string, number>();
const heldDeletions = new Set<{ fonts: FontFaceSet; face: FontFace; family: string }>();
const fontFamilyKey = (family: string) => family.trim().replace(/^["']|["']$/g, "");

/**
 * Keep the faces PDF.js loaded under these families (a CSS font-family list)
 * until the returned release is called: text a drawing sets in a PDF's own
 * face, by PDF.js's name for it, stays in that face while the page goes on
 * to another revision. A face PDF.js removes meanwhile leaves when the last
 * hold on its family is released; every other face leaves as before.
 */
export function holdPdfJsFontFaces(fontFamily: string | null | undefined): () => void {
  const families = [...new Set((fontFamily ?? "").split(",").map(fontFamilyKey).filter(Boolean))];
  for (const family of families) fontHolds.set(family, (fontHolds.get(family) ?? 0) + 1);
  let released = false;
  return () => {
    if (released) return;
    released = true;
    for (const family of families) {
      const left = (fontHolds.get(family) ?? 1) - 1;
      if (left > 0) { fontHolds.set(family, left); continue; }
      fontHolds.delete(family);
      for (const pending of [...heldDeletions]) {
        if (pending.family !== family) continue;
        heldDeletions.delete(pending);
        pending.fonts.delete(pending.face);
      }
    }
  };
}

const documents = new WeakMap<Document, Document>();

/** `doc` as PDF.js should see it: the same document, its fonts added as above. */
export function pdfJsFontDocument(doc: Document): Document {
  const known = documents.get(doc);
  if (known) return known;
  if (!doc.fonts || typeof FontFace === "undefined") return doc;
  const fonts = new DeferredFontSet(doc.fonts);
  const proxy = new Proxy(doc, {
    get(target, property) {
      if (property === "fonts") return fonts;
      const value = Reflect.get(target, property, target);
      return typeof value === "function" ? value.bind(target) : value;
    },
  });
  documents.set(doc, proxy);
  return proxy;
}
