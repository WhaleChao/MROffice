/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Browser print/PDF asset gate.
 *
 * A cloned print tree starts a new image lifecycle in the popup document.
 * Printing after fonts alone are ready races that lifecycle and intermittently
 * captures blank image boxes. This module owns the one bounded readiness
 * contract shared by every client-side document print path.
 */
import { layoutPending } from "@mrnote/core/web/components/console/editor/native-document/paged/layoutPending.ts";
import { officeTextOutlinesPending } from "@mrnote/core/web/lib/office-text-outline/office-text-outline-css.ts";

export const PRINT_ASSET_TIMEOUT_MS = 10_000;
/** How long the pages must hold still before they are copied for printing. */
const PRINT_SOURCE_QUIET_MS = 300;
/** How long Print waits out a display-font plan the page reports
 * unavailable: its two automatic retries, each with a 15 s deadline. */
export const PRINT_FONT_PLAN_RETRY_GRACE_MS = 45_000;
const PRINT_DIALOG_DELAY_MS = 100;

export type PrintReadinessErrorCode =
  | "aborted"
  | "font-load"
  | "image-load"
  | "image-decode"
  | "chart-render"
  | "timeout"
  | "print-failed"
  // The browser refused to open the print window (a pop-up blocker).
  | "popup-blocked";

/** What a "timeout" waited for: the document's own layout to settle (the
 * source), or the print copy's pictures and fonts (its assets). */
export type PrintReadinessStage = "layout" | "assets";

export class PrintReadinessError extends Error {
  readonly code: PrintReadinessErrorCode;
  readonly detail?: unknown;
  readonly stage?: PrintReadinessStage;

  constructor(
    code: PrintReadinessErrorCode,
    message: string,
    detail?: unknown,
    stage?: PrintReadinessStage,
  ) {
    super(message);
    this.name = "PrintReadinessError";
    this.code = code;
    this.detail = detail;
    if (stage) this.stage = stage;
  }
}

/** The reason a failed print tells the reader, one message each. */
export type PrintFailureReason =
  | "layout"
  | "timeout"
  | "font"
  | "image-load"
  | "image-decode"
  | "chart"
  | "changed"
  | "other";

/** The message for each reason; a reason without one shows the caller's
 * single message. */
export type PrintFailureMessages = Partial<Record<PrintFailureReason, string>>;

/** The reason of a print failure, from its code (and, for a timeout, what it
 * waited for). */
export function printFailureReason(error: PrintReadinessError): PrintFailureReason {
  switch (error.code) {
    case "timeout":
      return error.stage === "layout" ? "layout" : "timeout";
    case "font-load":
      return "font";
    case "image-load":
      return "image-load";
    case "image-decode":
      return "image-decode";
    case "chart-render":
      return "chart";
    case "aborted":
      return "changed";
    default:
      return "other";
  }
}

/** The console-notebooks message keys of the failure reasons. */
export type PrintFailureMessageKey =
  | "toolbar.printFailedLayout"
  | "toolbar.printFailedTimeout"
  | "toolbar.printFailedFont"
  | "toolbar.printFailedImageLoad"
  | "toolbar.printFailedImageDecode"
  | "toolbar.printFailedChart"
  | "toolbar.printFailedChanged"
  | "toolbar.printFailedOther";

/** What stopped a print, in the reader's language: one message per reason
 * (printFailureReason), not one sentence for every cause. */
export function documentPrintFailureMessages(t: (key: PrintFailureMessageKey) => string): PrintFailureMessages {
  return {
    layout: t("toolbar.printFailedLayout"),
    timeout: t("toolbar.printFailedTimeout"),
    font: t("toolbar.printFailedFont"),
    "image-load": t("toolbar.printFailedImageLoad"),
    "image-decode": t("toolbar.printFailedImageDecode"),
    chart: t("toolbar.printFailedChart"),
    changed: t("toolbar.printFailedChanged"),
    other: t("toolbar.printFailedOther"),
  };
}

/** The message a failed print shows: its reason's, unless
 * `window.__docsPrintFailureReasonOff = true` (A/B) or none was given. */
export function printFailureMessage(
  error: PrintReadinessError | undefined,
  fallback: string,
  messages?: PrintFailureMessages,
  win: Window | null = typeof window === "undefined" ? null : window,
): string {
  if (!error || !messages) return fallback;
  if ((win as (Window & { __docsPrintFailureReasonOff?: boolean }) | null)?.__docsPrintFailureReasonOff === true) return fallback;
  return messages[printFailureReason(error)] ?? fallback;
}

export interface PrintReadinessOptions {
  /** One deadline for fonts plus every image in the cloned print tree. */
  timeoutMs?: number;
  /** Allows the caller to cancel a superseded/closed print preparation. */
  signal?: AbortSignal;
}

export interface PrintWindowReadinessOptions extends PrintReadinessOptions {
  errorMessage?: string;
  /** Each failure reason's own message (printFailureMessages). */
  errorMessages?: PrintFailureMessages;
  onError?: (error: PrintReadinessError) => void;
  printDelayMs?: number;
  /**
   * Stop once the print copy is ready, without the print dialog: the server
   * renders the document's PDF from this very copy (a headless browser
   * waits for `data-mrai-print-ready` and prints it to a file).
   */
  skipDialog?: boolean;
}

/** On the print window's root once its copy is ready (`skipDialog`). */
export const PRINT_READY_ATTRIBUTE = "data-mrai-print-ready";
/** On the print window's root: the physical pages its copy prints. */
export const PRINT_PAGES_ATTRIBUTE = "data-mrai-print-pages";

export interface PrintWindowResult {
  status: "printed" | "ready" | "closed" | "failed";
  error?: PrintReadinessError;
}

function abortedError(): PrintReadinessError {
  return new PrintReadinessError("aborted", "Print preparation was cancelled.");
}

function withAbort<T>(
  promise: PromiseLike<T>,
  signal: AbortSignal,
): Promise<T> {
  if (signal.aborted) return Promise.reject(abortedError());
  return new Promise<T>((resolve, reject) => {
    const abort = () => {
      cleanup();
      reject(abortedError());
    };
    const cleanup = () => signal.removeEventListener("abort", abort);
    signal.addEventListener("abort", abort, { once: true });
    Promise.resolve(promise).then(
      (value) => {
        cleanup();
        resolve(value);
      },
      (error: unknown) => {
        cleanup();
        reject(error);
      },
    );
  });
}

/** Wait on the live source, before cloning. A copied pending flag cannot settle. */
export function documentPrintSourceReadiness(root: HTMLElement): "ready" | "pending" | "unavailable" {
  if (root.getAttribute("data-document-display-fonts") === "unavailable") return "unavailable";
  // A face still loading lays text out again when it arrives (the tabs, the
  // line boxes measured from it): Print pressed while a page just opened
  
  if ((root.ownerDocument.fonts as FontFaceSet | undefined)?.status === "loading" || layoutPending(root)) return "pending";
  // WordArt outlines still to be drawn over their words (office-text-outline-overlay.ts).
  if (officeTextOutlinesPending()) return "pending";
  return root.matches("[data-docs-exact-line-pending],[data-docs-auto-line-spacing-pending],[data-docs-display-fonts-pending],[data-docs-display-fonts-settling],[data-pagination-initial-settling],[data-pagination-background],[data-paged-bulk-filling],[data-pagination-stalled],[data-docs-page-bands-pending],[data-docs-word-cjk-spacing-pending]") ? "pending" : "ready";
}

/** How the print source came to be copied: settled, or taken as it stands
 * when the wait ran out with every block placed (`printUnsettled`). */
export type PrintSourceState = "settled" | "unsettled";

/** `window.__docsPrintUnsettledOff = true` (A/B): a source that has not
 * settled by the deadline fails the print, as before. */
function printUnsettledOff(): boolean {
  return typeof window !== "undefined" && (window as Window & { __docsPrintUnsettledOff?: boolean }).__docsPrintUnsettledOff === true;
}

/** Markers that say content is still to be placed on pages (not only laid
 * out again): a copy taken now would miss some of it. */
const UNPLACED_CONTENT = "[data-pagination-initial-settling],[data-pagination-background],[data-paged-bulk-filling],[data-pagination-stalled]";

/**
 * Whether the pages as they stand are a whole copy of the document: every
 * block placed (no marker of content still to place, no hollow sheet, and
 * every top-level block of the document counted on a sheet, its split pieces
 * as one), no sheet's content past its paper's foot, and no block on two
 * sheets but as the pieces of one split. A copy that fails any of these is
 * not printed.
 */
export function printSourceWhole(root: HTMLElement): boolean {
  if (root.matches(UNPLACED_CONTENT)) return false;
  const sheets = Array.from(root.querySelectorAll<HTMLElement>(":scope > article[data-page-sheet]"));
  if (sheets.length === 0) return false;
  const seen = new Set<string>();
  for (const sheet of sheets) {
    if (sheet.hasAttribute("data-page-hollow")) return false;
    const rect = sheet.getBoundingClientRect();
    const style = getComputedStyle(sheet);
    const paper = Number.parseFloat(style.getPropertyValue("--page-height-px")) || rect.height;
    const foot = rect.top + Math.min(rect.height, paper) - (Number.parseFloat(style.paddingBottom) || 0);
    for (const child of Array.from(sheet.children) as HTMLElement[]) {
      if (child.matches("[data-float], [data-float-mode], .note-editor-image-resizable, [data-sheet-print-part]")) continue;
      const box = child.getBoundingClientRect();
      if (box.height > 0 && box.bottom - foot > 2) return false;
      const node = (child as HTMLElement & { pmViewDesc?: ViewDescLike }).pmViewDesc?.node;
      const attrs = node?.attrs;
      if (!attrs || attrs.pageSplitPieceOf || child.getAttribute("data-page-split-continuation") === "true") continue;
      const id = typeof attrs.block_id === "string" ? attrs.block_id : null;
      if (!id) continue;
      if (seen.has(id)) return false;
      seen.add(id);
    }
  }
  // Nor a block missing: every top-level block of the document (its split
  // pieces counted once) stands on a sheet.
  const drawn = new Set<string>();
  for (const sheet of sheets) {
    for (const child of Array.from(sheet.children)) {
      const id = blockRootId((child as HTMLElement & { pmViewDesc?: ViewDescLike }).pmViewDesc?.node);
      if (id) drawn.add(id);
    }
  }
  const doc = (root as HTMLElement & { pmViewDesc?: ViewDescLike }).pmViewDesc?.node;
  if (doc?.forEach) {
    let missing = false;
    doc.forEach((child) => {
      const blocks = child.type?.name === "pageBlock" ? child : null;
      const visit = (block: NodeLike) => {
        const id = blockRootId(block);
        if (id && !drawn.has(id)) missing = true;
      };
      if (blocks?.forEach) blocks.forEach(visit);
      else visit(child);
    });
    if (missing) return false;
  }
  return true;
}

interface NodeLike {
  attrs?: Record<string, unknown>;
  type?: { name?: string };
  forEach?: (callback: (child: NodeLike) => void) => void;
}
interface ViewDescLike { node?: NodeLike }

/** A top-level block's identity, its split pieces counted as the block. */
function blockRootId(node: NodeLike | undefined): string | null {
  const attrs = node?.attrs;
  if (!attrs) return null;
  const piece = attrs.pageSplitPieceOf;
  if (typeof piece === "string" && piece) return piece;
  return typeof attrs.block_id === "string" && attrs.block_id ? attrs.block_id : null;
}

/** After the layout's next commit: two frames, so a dispatch under way has
 * drawn its pages before they are read. */
function afterNextCommit(win: Window | null): Promise<void> {
  if (!win || typeof win.requestAnimationFrame !== "function") return new Promise((resolve) => setTimeout(resolve, 32));
  return new Promise((resolve) => win.requestAnimationFrame(() => win.requestAnimationFrame(() => resolve())));
}

export function waitForDocumentPrintSource(
  root: HTMLElement,
  options: PrintReadinessOptions & {
    cancelled?: () => boolean;
    /** At the deadline, a source whose blocks are all placed is copied as
     * it stands ("unsettled") instead of failing: a page whose layout never
     * settles still prints, as Word prints what it shows. */
    printUnsettled?: boolean;
  } = {},
): Promise<PrintSourceState> {
  const deadline = Date.now() + (options.timeoutMs ?? PRINT_ASSET_TIMEOUT_MS);
  const connected = root.isConnected;
  return new Promise((resolve, reject) => {
    let timer: ReturnType<typeof setTimeout> | undefined;
    let readySamples = 0;
    let unavailableSince: number | null = null;
    // The pages still being drawn (a text box's tab stops sized after its
    
    let lastChange = Date.now();
    const changes = typeof MutationObserver === "undefined" ? null : new MutationObserver(() => {
      lastChange = Date.now();
    });
    changes?.observe(root, { subtree: true, childList: true, attributes: true, characterData: true });
    const finish = (error?: PrintReadinessError, state: PrintSourceState = "settled") => {
      changes?.disconnect();
      clearTimeout(timer); options.signal?.removeEventListener("abort", abort);
      if (error) reject(error); else resolve(state);
    };
    // The deadline with every block placed: the pages as the next commit
    // leaves them, if they are a whole copy (printSourceWhole).
    const asItStands = () => {
      const timedOut = new PrintReadinessError("timeout", "Document layout did not settle before printing.", undefined, "layout");
      if (!options.printUnsettled || printUnsettledOff() || root.matches(UNPLACED_CONTENT)) return finish(timedOut);
      void afterNextCommit(root.ownerDocument.defaultView).then(() => {
        if (options.signal?.aborted || options.cancelled?.() || (connected && !root.isConnected)) return abort();
        if (printSourceWhole(root)) finish(undefined, "unsettled");
        else finish(timedOut);
      });
    };
    const abort = () => finish(abortedError());
    const check = () => {
      if (options.signal?.aborted || options.cancelled?.() || (connected && !root.isConnected)) return abort();
      const state = documentPrintSourceReadiness(root);
      // A font plan that missed its deadline is asked for again on its own
      // (useDocumentDisplayFonts: twice, 3 s and 6 s later) while the page
      // says "unavailable": only a refusal that outlasts those retries ends
      // the print. Ended at once, Print on a page still opening in a busy
      
      // was pressed) said its pictures or fonts had failed, and the browser's
      // own print of the app that followed gave 10 pages for 5.
      const now = Date.now();
      if (state === "unavailable") {
        unavailableSince ??= now;
        if (now - unavailableSince >= PRINT_FONT_PLAN_RETRY_GRACE_MS || now >= deadline) {
          return finish(new PrintReadinessError("font-load", "Document display fonts are unavailable."));
        }
      } else unavailableSince = null;
      if (state === "ready" && ++readySamples >= 2 && now - lastChange >= PRINT_SOURCE_QUIET_MS) return finish();
      if (state !== "ready") readySamples = 0;
      if (now >= deadline) return asItStands();
      timer = setTimeout(check, 32);
    };
    options.signal?.addEventListener("abort", abort, {once:true});
    check();
  });
}

function waitForImageLoad(
  image: HTMLImageElement,
  signal: AbortSignal,
): Promise<void> {
  if (signal.aborted) return Promise.reject(abortedError());
  if (image.complete) return Promise.resolve();

  return new Promise<void>((resolve, reject) => {
    const cleanup = () => {
      image.removeEventListener("load", loaded);
      image.removeEventListener("error", failed);
      signal.removeEventListener("abort", aborted);
    };
    const loaded = () => {
      cleanup();
      resolve();
    };
    const failed = (event: Event) => {
      cleanup();
      reject(
        new PrintReadinessError(
          "image-load",
          "An image in the print document failed to load.",
          event,
        ),
      );
    };
    const aborted = () => {
      cleanup();
      reject(abortedError());
    };
    image.addEventListener("load", loaded, { once: true });
    image.addEventListener("error", failed, { once: true });
    signal.addEventListener("abort", aborted, { once: true });

    // Close the small race where the resource completed between the initial
    // `complete` check and listener registration.
    if (image.complete) loaded();
  });
}

async function waitForImage(
  image: HTMLImageElement,
  signal: AbortSignal,
): Promise<void> {
  await waitForImageLoad(image, signal);
  if (!image.complete || image.naturalWidth <= 0) {
    throw new PrintReadinessError(
      "image-load",
      "An image in the print document has no decodable pixels.",
    );
  }

  if (typeof image.decode !== "function") return;
  try {
    await withAbort(image.decode(), signal);
  } catch (error) {
    if (error instanceof PrintReadinessError) throw error;
    // Chromium refuses decode() for a picture too large to decode whole at
    // once (an EncodingError), though it paints it, scaled down, as ever:
    // R404-D02's 7000 x 5000 photographs failed every print of the
    // document, and its PDF came from the fallback writer (31 pages where
    // Word has 24). Its pixels are read again at a small size; only a
    // picture that cannot give even those fails the print.
    if (await decodesScaledDown(image, signal)) return;
    throw new PrintReadinessError(
      "image-decode",
      "An image in the print document failed to decode.",
      error,
    );
  }
}

/** `window.__docsPrintScaledDecodeOff = true`: a refused decode() fails the
 * print, as before (A/B of one build). */
function scaledDecodeOff(): boolean {
  return typeof window !== "undefined" &&
    (window as Window & { __docsPrintScaledDecodeOff?: boolean }).__docsPrintScaledDecodeOff === true;
}

/** Whether a loaded picture whose decode() was refused gives its pixels at a
 * small size (createImageBitmap's scaled decode). */
async function decodesScaledDown(image: HTMLImageElement, signal: AbortSignal): Promise<boolean> {
  if (scaledDecodeOff() || !(image.naturalWidth > 0) || !(image.naturalHeight > 0)) return false;
  const view = image.ownerDocument?.defaultView as (Window & typeof globalThis) | null;
  const scope = typeof view?.createImageBitmap === "function" ? view : typeof createImageBitmap === "function" ? globalThis : null;
  if (!scope) return false;
  try {
    const bitmap = await withAbort(
      scope.createImageBitmap(image, { resizeWidth: Math.min(64, image.naturalWidth), resizeQuality: "low" }),
      signal,
    );
    bitmap.close?.();
    return true;
  } catch (error) {
    if (error instanceof PrintReadinessError) throw error;
    return false;
  }
}

/**
 * Wait for the cloned document's fonts and all images under `root`.
 *
 * The timeout is deliberately global rather than per image: a document with
 * many broken/slow images cannot multiply the user's wait indefinitely.
 */
export async function waitForPrintReadiness(
  doc: Document,
  root: ParentNode = doc,
  options: PrintReadinessOptions = {},
): Promise<void> {
  const timeoutMs = options.timeoutMs ?? PRINT_ASSET_TIMEOUT_MS;
  if (!Number.isFinite(timeoutMs) || timeoutMs <= 0) {
    throw new RangeError("Print readiness timeout must be a positive number.");
  }

  const scoped = new AbortController();
  let timedOut = false;
  const timeout = globalThis.setTimeout(() => {
    timedOut = true;
    scoped.abort();
  }, timeoutMs);
  const externalAbort = () => scoped.abort();
  options.signal?.addEventListener("abort", externalAbort, { once: true });
  if (options.signal?.aborted) scoped.abort();

  try {
    const fonts = (doc as Document & { fonts?: FontFaceSet }).fonts;
    const fontReady = fonts?.ready
      ? withAbort(fonts.ready, scoped.signal).catch((error: unknown) => {
          if (error instanceof PrintReadinessError) throw error;
          throw new PrintReadinessError(
            "font-load",
            "The print document's fonts failed to load.",
            error,
          );
        })
      : Promise.resolve();
    // Only pictures are waited for. ProseMirror keeps the caret in place with
    // an `<img class="ProseMirror-separator" mark-placeholder>` that has no
    // source and no pixels; treating it as a picture that failed to load
    // refused every print of a document holding one (a paragraph mark, an
    // empty line). Any other image counts, a picture still without a source
    // included: it has nothing to print, so the print refuses visibly.
    const images = Array.from(root.querySelectorAll<HTMLImageElement>("img")).filter(
      (image) =>
        !image.classList.contains("ProseMirror-separator") &&
        !image.hasAttribute("mark-placeholder"),
    );
    // Off-screen document images are often marked lazy in the live editor.
    // The print popup needs every page, so explicitly start those requests.
    for (const image of images) image.loading = "eager";
    await Promise.all([
      fontReady,
      ...images.map((image) => waitForImage(image, scoped.signal)),
    ]);
  } catch (error) {
    if (timedOut) {
      throw new PrintReadinessError(
        "timeout",
        `Print assets did not become ready within ${timeoutMs} ms.`,
        error,
        "assets",
      );
    }
    if (options.signal?.aborted) throw abortedError();
    throw error;
  } finally {
    // Cancel sibling waits after the first failure as well as after success;
    // otherwise listeners on other slow images would outlive this operation.
    scoped.abort();
    globalThis.clearTimeout(timeout);
    options.signal?.removeEventListener("abort", externalAbort);
  }
}

const DEFAULT_FAILURE_MESSAGE =
  "Unable to prepare this document for printing because an image or font did not load. Close this window and try again.";

/** Replace the unprintable clone with an explicit, accessible failure. */
export function renderPrintReadinessFailure(
  doc: Document,
  message = DEFAULT_FAILURE_MESSAGE,
  error?: PrintReadinessError,
  /** Each reason's own message (printFailureMessages): the reader is told
   * what stopped the print, not one sentence for every cause. */
  messages?: PrintFailureMessages,
): void {
  const alert = doc.createElement("div");
  alert.setAttribute("role", "alert");
  alert.setAttribute("data-testid", "print-readiness-error");
  // What stopped the copy, for support and probes.
  if (error) {
    alert.setAttribute("data-print-failure", `${error.code}${error.stage ? `/${error.stage}` : ""}: ${error.message}`.slice(0, 300));
    alert.setAttribute("data-print-failure-reason", printFailureReason(error));
  }
  alert.setAttribute(
    "style",
    "max-width:42rem;margin:10vh auto;padding:1.25rem 1.5rem;border:1px solid #c2413b;border-radius:.75rem;background:#fff7f6;color:#7f1d1d;font:600 16px/1.6 system-ui,sans-serif",
  );
  alert.textContent = printFailureMessage(error, message, messages, doc.defaultView ?? (typeof window === "undefined" ? null : window));
  doc.body.replaceChildren(alert);
}

function normalizeReadinessError(error: unknown): PrintReadinessError {
  if (error instanceof PrintReadinessError) return error;
  return new PrintReadinessError(
    "print-failed",
    "Preparing the print dialog failed.",
    error,
  );
}

function waitForPrintDialogDelay(
  delayMs: number,
  signal?: AbortSignal,
): Promise<void> {
  if (!Number.isFinite(delayMs) || delayMs < 0) {
    return Promise.reject(
      new RangeError("Print dialog delay must be a non-negative number."),
    );
  }
  if (signal?.aborted) return Promise.reject(abortedError());
  return new Promise<void>((resolve, reject) => {
    const aborted = () => {
      globalThis.clearTimeout(timer);
      signal?.removeEventListener("abort", aborted);
      reject(abortedError());
    };
    // The opener realm owns this timer. A popup can be closed or have its
    // event loop suspended without stranding the caller's Promise forever.
    const timer = globalThis.setTimeout(() => {
      signal?.removeEventListener("abort", aborted);
      resolve();
    }, delayMs);
    signal?.addEventListener("abort", aborted, { once: true });
  });
}

/**
 * Shared terminal step for all client-side PDF/print paths. Asset failure is
 * fail-closed: `window.print()` is never invoked and the popup explains why.
 */
export async function printWindowWhenReady(
  printWindow: Window,
  root: ParentNode,
  options: PrintWindowReadinessOptions = {},
): Promise<PrintWindowResult> {
  const timeoutMs = options.timeoutMs ?? PRINT_ASSET_TIMEOUT_MS;
  const deadline = Date.now() + timeoutMs;
  try {
    if (!Number.isFinite(timeoutMs) || timeoutMs <= 0) {
      throw new RangeError(
        "Print readiness timeout must be a positive number.",
      );
    }
    await waitForPrintReadiness(printWindow.document, root, {
      timeoutMs,
      signal: options.signal,
    });
    if (printWindow.closed) return { status: "closed" };
    if (options.signal?.aborted) throw abortedError();
    if (options.skipDialog) {
      printWindow.document.documentElement.setAttribute(PRINT_READY_ATTRIBUTE, "true");
      return { status: "ready" };
    }
    const delayMs = options.printDelayMs ?? PRINT_DIALOG_DELAY_MS;
    if (delayMs > deadline - Date.now()) {
      throw new PrintReadinessError(
        "timeout",
        `Print assets and dialog did not become ready within ${timeoutMs} ms.`,
        undefined,
        "assets",
      );
    }
    await waitForPrintDialogDelay(delayMs, options.signal);
    if (printWindow.closed) return { status: "closed" };
    if (options.signal?.aborted) throw abortedError();
    printWindow.print();
    return { status: "printed" };
  } catch (rawError) {
    const error = normalizeReadinessError(rawError);
    if (!printWindow.closed) {
      renderPrintReadinessFailure(printWindow.document, options.errorMessage, error, options.errorMessages);
    }
    options.onError?.(error);
    return { status: "failed", error };
  }
}
