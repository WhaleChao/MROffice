/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * MR Office / MRNote Classic (docs/desktop-classic.md): the window's instant
 * view (apps/desktop src/classic/instant-view.ts) covers the window from the
 * moment it is made, before this page and its data are there: the home as it
 * was last shown, or a double-clicked deck's first slide. This page says when
 * it shows the same itself, so the cover goes:
 *
 *  - "home": the notebook's page with its files listed (or its empty list),
 *    Classic's library with its notebooks (or its empty state);
 *  - "deck": a deck's first slide painted by its editor (the editor says so:
 *    lib/presentation/first-slide-painted), or a dialog or an alert shown.
 *
 * And once the home is shown and the page is idle, it hands the shell the
 * home as it is (its HTML without scripts, frames or canvases; the app's own
 * style sheets and fonts are read from the build on disk), for the next
 * launch's cover.
 */
import { IS_DESKTOP_CLASSIC } from "./desktop-classic";
import { PRESENTATION_FIRST_SLIDE_PAINTED_EVENT } from "./presentation/first-slide-painted";

interface InstantBridge {
  instantDone?(kind: "home" | "deck"): Promise<unknown>;
  instantSave?(html: string): Promise<unknown>;
  /** The double-clicked document the cover shows (absent in older builds). */
  instantDocument?(): Promise<unknown>;
  /** A document opened while this page is shown covers the window (absent in older builds). */
  onInstantWatch?(listener: () => void): () => void;
  instantWatching?(): Promise<unknown>;
}

const SNAPSHOT_MAX_CHARS = 3_000_000;

function bridge(): InstantBridge | null {
  const desktop = (window as unknown as { mraiDesktop?: InstantBridge }).mraiDesktop;
  return desktop && typeof desktop.instantDone === "function" ? desktop : null;
}

/** The home is shown with its data (or its honest empty state). */
export function homeShown(doc: Document = document): boolean {
  const notebook =
    doc.querySelector('[data-testid="notebook-canvas"]') &&
    (doc.querySelector('[data-testid="sidebar-page-owner-group"]') ||
      doc.querySelector('[data-testid="notebook-sidebar-pages-empty"]'));
  const library =
    doc.querySelector('[data-testid="notebooks-create-button"]') &&
    !doc.querySelector('[data-testid="notebook-library-loading"]');
  return Boolean(notebook || library);
}

/**
 * The page asks the person something or tells them something went wrong (a
 * dialog, an alert): its cover must not hide that, whatever it stands for.
 */
export function pageSpeaks(doc: Document = document): boolean {
  for (const node of doc.querySelectorAll<HTMLElement>('[role="alertdialog"], [role="dialog"][aria-modal="true"], [role="alert"]')) {
    if ((node.textContent ?? "").trim() && node.getClientRects().length > 0) return true;
  }
  return false;
}

/** How long the page looks for what its cover shows (the shell gives a cover 45 s). */
/**
 * A Word document, a workbook or a PDF drawn in the front window: what an
 * instant view of its first page stands in for (a deck says so itself,
 * PRESENTATION_FIRST_SLIDE_PAINTED_EVENT). Without it the picture of a Word
 * document or a workbook that carried one stayed over the open document
 * until the shell's own limit, 45 s.
 */
export function documentShown(doc: Document = document, name: string | null = null): boolean {
  const front = doc.querySelector<HTMLElement>(".wm-window--focused");
  if (!front) return false;
  // The cover's own document (its window's title names it): an earlier
  // document's window brought back with the desk is not what it shows.
  if (name && !(front.querySelector('[data-testid="window-title-text"]')?.textContent ?? "").includes(documentStem(name))) {
    return false;
  }
  for (const paragraph of front.querySelectorAll(".ProseMirror p, .ProseMirror h1, .ProseMirror h2, .ProseMirror td")) {
    if ((paragraph.textContent ?? "").trim()) {
      // A Word document's pages are hidden while its first layout settles
      // (the page ledger's pagesHiddenForInitialSettle): its words are in
      // the page then, not on screen. The picture stays until they show --
      // taken off at the words, the window showed nothing of the document
      
      const root = paragraph.closest<HTMLElement>(".ProseMirror");
      return !root || !pagesSettling(root);
    }
  }
  const grid = front.querySelector('[data-testid="spreadsheet-grid-viewport"]');
  // A workbook's grid is there before its cells are drawn (up to 1 s, MR
  
  // with something painted.
  if (grid) return Boolean(grid.querySelector('[role="gridcell"]')) || [...front.querySelectorAll("canvas")].some(canvasPainted);
  const page = front.querySelector(".pdf-viewer__page:not(.off)");
  return Boolean(page?.querySelector("canvas") && page.querySelector(".pdf-viewer__text-layer span"));
}

/** Something drawn in the middle of the canvas's top quarter (where a grid's first rows are). */
function canvasPainted(canvas: HTMLCanvasElement): boolean {
  try {
    if (!canvas.width || !canvas.height) return false;
    const pixels = canvas.getContext("2d")?.getImageData(canvas.width >> 1, canvas.height >> 2, 40, 40).data;
    if (!pixels) return false;
    for (let at = 3; at < pixels.length; at += 4) if (pixels[at]) return true;
  } catch {
    return false;
  }
  return false;
}

/** A file name as a window's title shows it: without its extension. */
export function documentStem(name: string): string {
  const dot = name.lastIndexOf(".");
  return dot > 0 ? name.slice(0, dot) : name;
}

/**
 * How long a document drawn in the front window that is not the cover's
 * (another title) holds before it counts all the same: a title the page
 * writes otherwise must not leave the picture over the open document until
 * the shell's own limit.
 */
const OTHER_DOCUMENT_GRACE_MS = 4_000;

/** The first layout's attributes on an editor root (paged/pageRebalanceExtension
 * INITIAL_PAGINATION_SETTLING_ATTR, documentLayoutStatus INITIAL_SETTLE_SHOWN_ATTR):
 * settling and not yet shown = its pages hidden. Named here, not imported:
 * this module is in every desktop page and the editor's are not. */
function pagesSettling(root: HTMLElement): boolean {
  return root.getAttribute("data-pagination-initial-settling") === "true" && root.getAttribute("data-pagination-settle-shown") !== "true";
}

const WATCH_MS = 50_000;
const GRID_DRAWN_AFTER_MS = 3_000;

const DROPPED =
  "script,noscript,iframe,canvas,video,audio,object,embed,template," +
  "link[rel='preload'],link[rel='modulepreload'],link[rel='prefetch'],link[rel='manifest'],meta[http-equiv],base";
/** Pictures the cover can show: the build's and the public folder's (never the API's). */
const KEPT_PICTURE = /^(?:data:|\/(?:_next\/static|fonts|icons|images|brand)\/)/;

/** The page as it is, for a cover that runs no script and has no server. */
export function homeSnapshot(doc: Document = document, origin = window.location.origin): string | null {
  const root = doc.documentElement.cloneNode(true) as HTMLElement;
  root.querySelectorAll(DROPPED).forEach((node) => node.remove());
  root.querySelectorAll("img").forEach((img) => {
    const src = (img.getAttribute("src") ?? "").replace(origin, "");
    if (!KEPT_PICTURE.test(src)) {
      img.removeAttribute("src");
      img.removeAttribute("srcset");
    }
  });
  const still = doc.createElement("style");
  still.textContent =
    "*,*::before,*::after{animation:none!important;transition:none!important;caret-color:transparent!important}html{pointer-events:none}";
  root.querySelector("head")?.appendChild(still);
  const html = `<!doctype html>${root.outerHTML}`.split(origin).join("");
  return html.length > SNAPSHOT_MAX_CHARS ? null : html;
}

let installed = false;

export function installDesktopInstant(): void {
  if (!IS_DESKTOP_CLASSIC || installed || typeof window === "undefined" || typeof MutationObserver === "undefined") return;
  const desktop = bridge();
  if (!desktop) return;
  installed = true;
  let home = false;
  let deck = false;
  const deckDone = () => {
    if (deck) return;
    deck = true;
    void desktop.instantDone?.("deck")?.catch(() => undefined);
  };
  // The editor painted a deck's first slide (lib/presentation/first-slide-painted).
  window.addEventListener(PRESENTATION_FIRST_SLIDE_PAINTED_EVENT, () => deckDone(), { once: true });
  const save = () => {
    if (!homeShown()) return;
    const html = homeSnapshot();
    if (html) void desktop.instantSave?.(html)?.catch(() => undefined);
  };
  // The home as the person leaves it (the window closed, the app quit), for
  // the next launch; the one taken once it first settled stands otherwise.
  window.addEventListener("beforeunload", () => {
    if (home) save();
  });
  const check = () => {
    if (!home && homeShown()) {
      home = true;
      void desktop.instantDone?.("home")?.catch(() => undefined);
      // Once the page has settled: the home as the person sees it.
      window.setTimeout(() => {
        if (typeof window.requestIdleCallback === "function") window.requestIdleCallback(save, { timeout: 5_000 });
        else save();
      }, 2_500);
    }
    if (!deck) {
      if (pageSpeaks() || documentShown(document, coverDocument)) deckDone();
      else if (document.querySelector('.wm-window--focused [data-testid="spreadsheet-grid-viewport"]')) {
        // A grid whose canvas cannot be read (or draws elsewhere) counts as
        // drawn 3 s after it came: never a cover left over a workbook.
        gridSeenAt ||= Date.now();
        if (Date.now() - gridSeenAt > GRID_DRAWN_AFTER_MS) deckDone();
        else if (!gridRecheck) gridRecheck = window.setTimeout(() => {
          gridRecheck = 0;
          check();
        }, 100);
      }
      else if (coverDocument && documentShown() && otherSince === null) {
        otherSince = window.setTimeout(() => {
          if (!deck && documentShown()) deckDone();
          otherSince = null;
        }, OTHER_DOCUMENT_GRACE_MS);
      }
    }
    if (home && deck) observer.disconnect();
  };
  let gridRecheck = 0;
  let gridSeenAt = 0;
  let coverDocument: string | null = null;
  let otherSince: number | null = null;
  const askCoverDocument = () =>
    void desktop
      .instantDocument?.()
      ?.then((name) => {
        if (typeof name === "string" && name) coverDocument = name;
      })
      .catch(() => undefined);
  askCoverDocument();
  // A document opened while this page is shown (the spare home of a start in
  // the background, a window in use) covers the window with its first page
  // (apps/desktop src/classic/main.ts coverForDocument): watched for again,
  // the cover told so (it then waits as long as at a launch).
  desktop.onInstantWatch?.(() => {
    deck = false;
    gridSeenAt = 0;
    coverDocument = null;
    if (otherSince !== null) window.clearTimeout(otherSince);
    otherSince = null;
    askCoverDocument();
    void desktop.instantWatching?.()?.catch(() => undefined);
    window.addEventListener(PRESENTATION_FIRST_SLIDE_PAINTED_EVENT, () => deckDone(), { once: true });
    observer.observe(document.documentElement, { childList: true, subtree: true });
    window.clearTimeout(watchEnd);
    watchEnd = window.setTimeout(() => observer.disconnect(), WATCH_MS);
    check();
  });
  let scheduled = false;
  const observer = new MutationObserver(() => {
    if (scheduled) return;
    scheduled = true;
    window.setTimeout(() => {
      scheduled = false;
      check();
    }, 50);
  });
  observer.observe(document.documentElement, { childList: true, subtree: true });
  // A cover is gone by then anyway (the shell's own limit, 45 s): no more
  // looking at every change of the page.
  let watchEnd = window.setTimeout(() => observer.disconnect(), WATCH_MS);
  check();
}
