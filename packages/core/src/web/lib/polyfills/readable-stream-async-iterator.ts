/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Safari / older WebKit do not implement async iteration of a `ReadableStream`
 * (`ReadableStream.prototype[Symbol.asyncIterator]` / `.values()` are missing,
 * tracked in WebKit bug 194379). PDF.js v5's `PDFPageProxy.getTextContent()`
 * does `for await (const value of readableStream)` (pdfjs-dist build/pdf.mjs:
 * ~15294), so on Safari every text-extraction path throws
 *   "undefined is not a function (near '...value of readableStream...')"
 * — which surfaced as a fatal modal right after a PDF upload (the editable
 * projection compiler and the in-viewer search both call `getTextContent`).
 *
 * This installs the WHATWG-spec async iterator (same shape the platforms that
 * DO support it expose), so `for await ... of` a stream works on Safari too.
 * Idempotent + feature-gated: a no-op on engines that already have it (Chrome,
 * Firefox, recent Safari), so it can't regress conformant browsers.
 *
 * Import this for its side effect BEFORE any code that async-iterates a
 * ReadableStream (i.e. before pdf.js is loaded).
 */

type AsyncIteratorOptions = { preventCancel?: boolean };

function installReadableStreamAsyncIterator(): void {
  if (typeof ReadableStream === "undefined") return;
  const proto = ReadableStream.prototype as unknown as Record<PropertyKey, unknown>;
  // Already supported (native async iteration) — leave it untouched.
  if (typeof proto[Symbol.asyncIterator] === "function") return;

  function values(
    this: ReadableStream,
    { preventCancel = false }: AsyncIteratorOptions = {},
  ): AsyncIterableIterator<unknown> {
    const reader = this.getReader();
    return {
      next() {
        return reader.read() as Promise<IteratorResult<unknown>>;
      },
      async return(value?: unknown) {
        if (!preventCancel) {
          const cancelPromise = reader.cancel(value);
          reader.releaseLock();
          await cancelPromise;
        } else {
          reader.releaseLock();
        }
        return { done: true, value } as IteratorResult<unknown>;
      },
      [Symbol.asyncIterator]() {
        return this;
      },
    };
  }

  // `.values()` is the spec method; `[Symbol.asyncIterator]` aliases it so
  // `for await ... of stream` resolves. Define both, non-enumerable.
  Object.defineProperty(proto, "values", {
    value: values,
    writable: true,
    configurable: true,
  });
  Object.defineProperty(proto, Symbol.asyncIterator, {
    value: values,
    writable: true,
    configurable: true,
  });
}

installReadableStreamAsyncIterator();

export {};
