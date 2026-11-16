/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Resource-lifecycle helpers for the PDF reader.
 *
 * These helpers only orchestrate public pdf.js APIs (`RenderTask.cancel`,
 * `TextLayer.cancel`, `PDFPageProxy.cleanup`, and
 * `PDFDocumentLoadingTask.destroy`). No pdf.js source is copied here.
 */

export interface PdfCancelableTask {
  cancel(extraDelay?: number): void;
}

export interface PdfCleanupPage {
  cleanup(resetStats?: boolean): boolean;
}

export interface PdfDestroyable {
  destroy(): void | Promise<void>;
}

export interface PdfCanvasBackingStore {
  width: number;
  height: number;
}

export type PdfTaskRegistry = Map<number, PdfCancelableTask>;

export interface ReleasedCanvasBackingStore {
  width: number;
  height: number;
  estimatedBytes: number;
}

/**
 * Cancel one tracked task and remove only the exact task that was observed.
 * The identity guard protects a replacement task installed synchronously by a
 * completion callback.
 */
export function cancelTrackedPdfTask(
  tasks: PdfTaskRegistry,
  pageNumber: number,
): boolean {
  const task = tasks.get(pageNumber);
  if (!task) return false;
  try {
    task.cancel();
  } catch {
    // A task that settled between lookup and cancellation is already safe.
  }
  if (tasks.get(pageNumber) === task) {
    tasks.delete(pageNumber);
  }
  return true;
}

/** Cancel every task currently owned by a registry. */
export function cancelTrackedPdfTasks(tasks: PdfTaskRegistry): number[] {
  const pageNumbers = [...tasks.keys()];
  for (const pageNumber of pageNumbers) {
    cancelTrackedPdfTask(tasks, pageNumber);
  }
  return pageNumbers;
}

/**
 * Build the unique set of rendered or in-flight pages outside the keep set.
 * Invalid persisted page numbers are ignored defensively.
 */
export function collectPdfPagesOutsideKeepSet(
  keepSet: ReadonlySet<number>,
  ...pageSources: ReadonlyArray<Iterable<number>>
): number[] {
  const release = new Set<number>();
  for (const source of pageSources) {
    for (const pageNumber of source) {
      if (
        Number.isInteger(pageNumber) &&
        pageNumber >= 1 &&
        !keepSet.has(pageNumber)
      ) {
        release.add(pageNumber);
      }
    }
  }
  return [...release].sort((a, b) => a - b);
}

/**
 * Release a canvas backing store and report its former approximate RGBA cost.
 * Resetting either dimension clears the bitmap in browsers; setting both keeps
 * the postcondition explicit and matches the viewer's existing cache reset.
 */
export function releasePdfCanvasBackingStore(
  canvas: PdfCanvasBackingStore,
): ReleasedCanvasBackingStore {
  const width =
    Number.isFinite(canvas.width) && canvas.width > 0
      ? Math.floor(canvas.width)
      : 0;
  const height =
    Number.isFinite(canvas.height) && canvas.height > 0
      ? Math.floor(canvas.height)
      : 0;
  const estimatedBytes = Math.min(
    Number.MAX_SAFE_INTEGER,
    width * height * 4,
  );
  canvas.width = 0;
  canvas.height = 0;
  return { width, height, estimatedBytes };
}

/**
 * Ask pdf.js to release page-level operator/image resources after the owning
 * render pipeline settles. Rejected/cancelled pipelines are expected and do
 * not suppress cleanup. `PDFPageProxy.cleanup()` itself safely returns false
 * when a newer render is still using the page.
 */
export async function cleanupPdfPageAfter(
  page: PdfCleanupPage,
  blocker?: PromiseLike<unknown> | null,
): Promise<boolean> {
  if (blocker) {
    try {
      await blocker;
    } catch {
      // Cancellation is precisely when cleanup is most useful.
    }
  }
  try {
    return page.cleanup();
  } catch {
    return false;
  }
}

/**
 * Create one idempotent destroy operation for a pdf.js document session.
 * The loading task owns the worker and network transport; a resolved document's
 * `destroy()` delegates to that same loading task, so only one owner is called.
 */
export function createPdfSessionDestroyer(
  resources: () => {
    loadingTask?: PdfDestroyable | null;
    document?: PdfDestroyable | null;
  },
): () => Promise<void> {
  let destroyPromise: Promise<void> | null = null;
  return () => {
    if (destroyPromise) return destroyPromise;
    const current = resources();
    const owner = current.loadingTask ?? current.document;
    if (!owner) return Promise.resolve();
    try {
      destroyPromise = Promise.resolve(owner.destroy()).then(() => undefined);
    } catch (error) {
      destroyPromise = Promise.reject(error);
    }
    return destroyPromise;
  };
}

/** The page renders of one document shown by the reader, from its
 * publication to the end of its pdf.js session. */
export interface PdfRenderSession {
  /**
   * `render`'s own outcome while the session lasts. Once it has ended, a
   * render still in flight is waited for no longer (resolved), and one asked
   * for afterwards is not begun.
   */
  bound(render: () => Promise<void>): Promise<void>;
  /** The document's session is over (its worker terminated). Idempotent. */
  end(): void;
}












export function createPdfRenderSession(): PdfRenderSession {
  let ended = false;
  const waiting = new Set<() => void>();
  return {
    bound(render) {
      if (ended) return Promise.resolve();
      return new Promise<void>((resolve, reject) => {
        const release = () => resolve();
        waiting.add(release);
        let rendering: Promise<void>;
        try {
          rendering = render();
        } catch (error) {
          rendering = Promise.reject(error);
        }
        rendering.then(
          () => {
            waiting.delete(release);
            resolve();
          },
          (error: unknown) => {
            waiting.delete(release);
            reject(error);
          },
        );
      });
    },
    end() {
      if (ended) return;
      ended = true;
      const released = [...waiting];
      waiting.clear();
      for (const release of released) release();
    },
  };
}

/** Classify the cancellation errors emitted by pdf.js canvas/text renderers. */
export function isPdfRenderCancellation(error: unknown): boolean {
  if (!error || typeof error !== "object") return false;
  const name =
    typeof (error as { name?: unknown }).name === "string"
      ? (error as { name: string }).name
      : "";
  if (name === "RenderingCancelledException" || name === "AbortException") {
    return true;
  }
  const message =
    typeof (error as { message?: unknown }).message === "string"
      ? (error as { message: string }).message
      : "";
  return /\brender(?:ing)? cancelled\b|\btext layer (?:task )?cancelled\b/i.test(
    message,
  );
}
