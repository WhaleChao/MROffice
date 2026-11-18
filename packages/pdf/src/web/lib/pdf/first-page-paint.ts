/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** The longest the other pages wait for the opening page's first paint. */
export const OPENING_PAGE_PAINT_MAX_MS = 4_000;

// Opening pages being drawn in this tab right now (any viewer), for the
// drawings that yield to them: a file's local preview, drawn by pdf.js on
// the same thread while the server made its page, stood over the editor's
// loading card and went on drawing its next pages while the editor drew the
// one the reader would edit (LocalFilePreviewLayer).
let openingsDrawing = 0;
const afterOpenings = new Set<() => void>();
const openingStarts = new Set<() => void>();
const openingEnds = new Set<() => void>();

/** Whether a viewer is drawing a document's opening page right now. */
export function pdfOpeningPaintActive(): boolean {
  return openingsDrawing > 0;
}

/** Settles once no viewer is drawing an opening page (at once when none is). */
export function afterPdfOpeningPaints(): Promise<void> {
  if (openingsDrawing === 0) return Promise.resolve();
  return new Promise((resolve) => {
    afterOpenings.add(resolve);
  });
}

/** Told each time a viewer starts drawing an opening page. */
export function onPdfOpeningPaintStart(listener: () => void): () => void {
  openingStarts.add(listener);
  return () => {
    openingStarts.delete(listener);
  };
}

/** Told each time no viewer is drawing an opening page any more. */
export function onPdfOpeningPaintEnd(listener: () => void): () => void {
  openingEnds.add(listener);
  return () => {
    openingEnds.delete(listener);
  };
}

function openingStarted(): void {
  openingsDrawing += 1;
  for (const listener of [...openingStarts]) listener();
}

function openingEnded(): void {
  openingsDrawing = Math.max(0, openingsDrawing - 1);
  if (openingsDrawing > 0) return;
  const waiting = [...afterOpenings];
  afterOpenings.clear();
  for (const resolve of waiting) resolve();
  for (const listener of [...openingEnds]) listener();
}

export interface OpeningPagePaint {
  readonly doc: object;
  /** The opening page, once its render has been asked for. */
  page: number | null;
  readonly done: Promise<void>;
  released: boolean;
  release(): void;
}

export interface OpeningPagePaintTurn {
  /** Start this page's render after this (the opening page's first paint). */
  readonly waitFor?: Promise<void>;
  /** This render is the opening page's: hand it its settled promise. */
  readonly opens?: (settled: Promise<void>) => void;
}

function armed(doc: object): OpeningPagePaint {
  let resolve: () => void = () => undefined;
  const done = new Promise<void>((settle) => {
    resolve = settle;
  });
  let timer: ReturnType<typeof setTimeout> | null = null;
  const gate: OpeningPagePaint = {
    doc,
    page: null,
    done,
    released: false,
    release() {
      if (gate.released) return;
      gate.released = true;
      if (timer !== null) clearTimeout(timer);
      timer = null;
      if (gate.page !== null) openingEnded();
      resolve();
    },
  };
  timer = setTimeout(() => gate.release(), OPENING_PAGE_PAINT_MAX_MS);
  return gate;
}

/**
 * The turn of a render of `pageNumber` of `doc`, asked while the reader is on
 * `currentPage`. The first render asked of a document arms its hold.
 */
export function openingPagePaint(
  ref: { current: OpeningPagePaint | null },
  doc: object,
  currentPage: number,
  pageNumber: number,
  densityOnly: boolean,
): OpeningPagePaintTurn | null {
  let gate = ref.current;
  if (!gate || gate.doc !== doc) {
    gate?.release();
    gate = armed(doc);
    ref.current = gate;
  }
  if (gate.released || densityOnly) return null;
  if (gate.page === null) {
    // A page asked before the opening page (one scrolled to, an observer
    // reporting first) goes as before: the hold is only for the pages asked
    // while the opening page is being drawn.
    if (pageNumber !== currentPage) return null;
    gate.page = pageNumber;
    openingStarted();
    const opening = gate;
    return {
      opens: (settled) => {
        void settled.then(() => opening.release(), () => opening.release());
      },
    };
  }
  if (pageNumber === gate.page) return null;
  // The reader is on another page than the one being drawn first: theirs
  // goes now.
  if (pageNumber === currentPage) {
    gate.release();
    return null;
  }
  return { waitFor: gate.done };
}
