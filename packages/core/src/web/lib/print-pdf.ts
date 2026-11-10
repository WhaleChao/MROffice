/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { offerOpenWhenAllowed } from "./download-file";

/**
 * Hand a finished PDF to the browser's print dialog.
 *
 * Excel for the web prints this way: it renders the sheet (print area, titles,
 * page breaks, scaling, orientation, margins, headers and footers) to a PDF
 * and opens the browser's print dialog for it, instead of printing the web
 * page around the grid. A hidden iframe keeps the user in the editor where the
 * browser's PDF viewer can print in place (Chromium, Edge, Firefox); where it
 * cannot, the PDF opens in a new tab whose viewer has its own Print button.
 */

export type PrintPdfOutcome = "dialog" | "tab" | "blocked";

/** Long enough for a print preview to spool; the next print replaces it sooner. */
const PRINT_FRAME_LIFETIME_MS = 10 * 60_000;
const PRINT_FRAME_LOAD_TIMEOUT_MS = 15_000;

let activeFrame: { frame: HTMLIFrameElement; url: string; timer: number } | null =
  null;

function releaseActiveFrame(): void {
  if (!activeFrame) return;
  window.clearTimeout(activeFrame.timer);
  activeFrame.frame.remove();
  URL.revokeObjectURL(activeFrame.url);
  activeFrame = null;
}

function openInTab(url: string, name?: string): PrintPdfOutcome {
  // No "noopener": with it window.open returns null even when the tab opens.
  const tab = window.open(url, "_blank");
  // The tab keeps the object URL alive only while this page lives; release it
  // once the viewer has surely read it.
  window.setTimeout(() => URL.revokeObjectURL(url), PRINT_FRAME_LIFETIME_MS);
  if (tab) return "tab";
  // The PDF arrived long after the click, and a pop-up no tap opened is
  // refused (iPad Safari always): the shell offers Open, a tap every
  // browser honours, instead of leaving the person with only an error.
  if (name && offerOpenWhenAllowed(name, () => { window.open(url, "_blank"); })) return "tab";
  return "blocked";
}

export interface PrintPdfOptions {
  /** The file's name, for the Open offer when the browser blocks the tab. */
  readonly name?: string;
}

export function printPdfBlob(blob: Blob, options: PrintPdfOptions = {}): Promise<PrintPdfOutcome> {
  releaseActiveFrame();
  const pdf =
    blob.type === "application/pdf"
      ? blob
      : new Blob([blob], { type: "application/pdf" });
  const url = URL.createObjectURL(pdf);
  const frame = document.createElement("iframe");
  frame.setAttribute("aria-hidden", "true");
  frame.tabIndex = -1;
  frame.title = "print";
  Object.assign(frame.style, {
    position: "fixed",
    right: "0",
    bottom: "0",
    width: "1px",
    height: "1px",
    border: "0",
    opacity: "0",
    pointerEvents: "none",
  });
  return new Promise<PrintPdfOutcome>((resolve) => {
    let settled = false;
    const settle = (outcome: PrintPdfOutcome) => {
      if (settled) return;
      settled = true;
      window.clearTimeout(loadTimer);
      resolve(outcome);
    };
    const fallBackToTab = () => {
      frame.remove();
      if (activeFrame?.frame === frame) {
        window.clearTimeout(activeFrame.timer);
        activeFrame = null;
      }
      settle(openInTab(url, options.name));
    };
    const loadTimer = window.setTimeout(fallBackToTab, PRINT_FRAME_LOAD_TIMEOUT_MS);
    frame.addEventListener(
      "load",
      () => {
        try {
          const view = frame.contentWindow;
          if (!view) throw new Error("no print frame");
          view.focus();
          view.print();
          settle("dialog");
        } catch {
          fallBackToTab();
        }
      },
      { once: true },
    );
    activeFrame = {
      frame,
      url,
      timer: window.setTimeout(releaseActiveFrame, PRINT_FRAME_LIFETIME_MS),
    };
    frame.src = url;
    document.body.appendChild(frame);
  });
}
