/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A large PDF read in ranges instead of downloaded whole.
 *
 * The reader normally fetches a PDF's bytes once and hands them to pdf.js:
 * the security fingerprint, the session cache and every pdf-lib edit read the
 * same bytes. A PDF past {@link PDF_RANGE_LOAD_MIN_BYTES} is not downloaded:
 * pdf.js asks for the ranges it needs (the cross-reference table first, then
 * the pages being shown) and they are fetched from the stored file's signed
 * link with HTTP Range. The link expires within the hour, so a refused range
 * asks for a fresh link and is fetched again.
 *
 * Such a document is read-only here (nothing that needs its whole bytes runs),
 * and pages open as quickly whatever the file's size.
 */

import type { PdfApi } from "../../components/notebook/contents/pdf/viewer/pdfApi";
import { PDF_EDITABLE_MAX_BYTES } from "@mrnote/core/web/lib/document-size-limits.ts";
import { fetchPrivateFile } from "@mrnote/core/web/lib/private-file-transfer.ts";

/** The size past which a PDF is read in ranges and opens read-only: the
 * editable PDF's document total (lib/document-size-limits). */
export const PDF_RANGE_LOAD_MIN_BYTES = PDF_EDITABLE_MAX_BYTES;
/**
 * The unit pdf.js reads in (its own default). It asks for consecutive missing
 * units as one range, so a page's 1 MB picture is still one request, while a
 * small object costs only its unit: opening a document that is not
 * linearized reads every page dictionary (pdf.js checks its last page by
 * walking the page tree), and where those sit beside each page's picture, as
 * a scanner writes them, a 1 MiB unit fetched the whole file.
 */
export const PDF_RANGE_CHUNK_BYTES = 64 * 1024;

export interface PdfRangeSource {
  readonly sizeBytes: number;
  /** A fresh signed link to the stored file (they expire). */
  readonly resolveUrl: () => Promise<string | null>;
}

export function shouldReadPdfInRanges(sizeBytes: number | null | undefined): boolean {
  return typeof sizeBytes === "number" && Number.isFinite(sizeBytes) && sizeBytes > PDF_RANGE_LOAD_MIN_BYTES;
}

const ATTEMPTS = 4;

class RangeFetcher {
  private url: string | null = null;
  private renewing: Promise<string | null> | null = null;

  constructor(
    private readonly source: PdfRangeSource,
    private readonly signal: AbortSignal,
  ) {}

  private async link(fresh: boolean): Promise<string> {
    if (!fresh && this.url) return this.url;
    if (!this.renewing) {
      this.renewing = this.source.resolveUrl().finally(() => {
        this.renewing = null;
      });
    }
    const url = await this.renewing;
    if (!url) throw new Error("pdf_range_link_unavailable");
    this.url = url;
    return url;
  }

  async fetch(begin: number, end: number): Promise<Uint8Array> {
    let lastError: unknown = null;
    for (let attempt = 0; attempt < ATTEMPTS; attempt += 1) {
      if (this.signal.aborted) throw new DOMException("Aborted", "AbortError");
      try {
        const url = await this.link(attempt > 0 && lastError instanceof RangeRefused);
        const response = await fetchPrivateFile(url, {
          headers: { Range: `bytes=${begin}-${end - 1}` },
          signal: this.signal,
          cache: "no-store",
        });
        if (response.status === 401 || response.status === 403 || response.status === 404) {
          throw new RangeRefused(response.status);
        }
        if (response.status === 206) {
          return new Uint8Array(await response.arrayBuffer());
        }
        if (response.ok) {
          // A store that ignores Range answers with the whole file; take the
          // range from it (only a small file reaches this, in practice).
          const whole = new Uint8Array(await response.arrayBuffer());
          return whole.slice(begin, end);
        }
        throw new Error(`pdf_range_status_${response.status}`);
      } catch (error) {
        if (error instanceof DOMException && error.name === "AbortError") throw error;
        lastError = error;
        await new Promise((resolve) => setTimeout(resolve, 400 * 2 ** attempt));
      }
    }
    throw lastError instanceof Error ? lastError : new Error("pdf_range_failed");
  }
}

class RangeRefused extends Error {
  constructor(readonly status: number) {
    super(`pdf_range_refused_${status}`);
  }
}

/**
 * The transport pdf.js reads ``source`` through. ``onFatal`` is told when a
 * range could not be fetched at all (pdf.js itself would wait forever).
 */
export function createPdfRangeTransport(
  pdfjs: PdfApi,
  source: PdfRangeSource,
  signal: AbortSignal,
  onFatal: (error: unknown) => void,
): InstanceType<PdfApi["PDFDataRangeTransport"]> {
  const fetcher = new RangeFetcher(source, signal);
  class SignedRangeTransport extends pdfjs.PDFDataRangeTransport {
    override requestDataRange(begin: number, end: number): void {
      void fetcher.fetch(begin, Math.min(end, source.sizeBytes)).then(
        (chunk) => {
          if (!signal.aborted) this.onDataRange(begin, chunk);
        },
        (error) => {
          if (!signal.aborted) onFatal(error);
        },
      );
    }
  }
  return new SignedRangeTransport(source.sizeBytes, null);
}
