/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Print a deck from the browser that shows it.
 *
 * Print used to save the deck, ask the server to print it to PDF in its own
 * headless browser (`render_presentation_pdf`, the page below), download that
 * PDF and hand it to the browser's print dialog: a save, a browser launch and
 * a whole PDF before the dialog opened. PowerPoint prints what is on screen.
 * The same print page (`/render/presentation-print`, the editor's own slide
 * renderer composed as PowerPoint's Print composes the slides) now opens in a
 * hidden frame of this page, is handed the deck as it is here, and the
 * browser prints it once every slide is drawn.
 *
 * The frame is this origin, so it draws each picture from the URL the editor
 * draws it from, with the reader's own access. When it cannot print -- no
 * frame, nothing drawn within {@link CLIENT_PRINT_DRAW_TIMEOUT_MS} -- the
 * answer is `"unavailable"` and the caller prints the server's PDF, as
 * before.
 */

export const CLIENT_PRINT_PATH = "/render/presentation-print";
/** The page asks for the deck with this message, and says it is drawn with the next. */
export const CLIENT_PRINT_WAITING = "mrai-deck-print:waiting";
export const CLIENT_PRINT_PAYLOAD = "mrai-deck-print:payload";
export const CLIENT_PRINT_DRAWN = "mrai-deck-print:drawn";
/** How long the frame may take to load and draw every slide. */
export const CLIENT_PRINT_DRAW_TIMEOUT_MS = 45_000;

export interface ClientPrintPayload {
  readonly content: Record<string, unknown>;
  readonly layout?: string;
  readonly include_hidden?: boolean;
  readonly embedded_fonts?: unknown;
  readonly now_ms?: number;
}

export type ClientPrintOutcome = "dialog" | "unavailable";

let active: HTMLIFrameElement | null = null;

function localePrefix(): string {
  const first = window.location.pathname.split("/")[1] ?? "";
  // The app's locale segment (`/zh/...`, `/pt-BR/...`); a path without one is
  // sent to the default locale by the app's middleware.
  return /^[a-z]{2}(?:-[A-Z]{2})?$/.test(first) ? `/${first}` : "";
}

/** Print `payload` from a hidden frame of this page; see the module doc. */
export function printDeckInBrowser(
  payload: ClientPrintPayload,
  options: { readonly timeoutMs?: number } = {},
): Promise<ClientPrintOutcome> {
  if (typeof window === "undefined" || typeof document === "undefined") return Promise.resolve("unavailable");
  active?.remove();
  const frame = document.createElement("iframe");
  active = frame;
  frame.setAttribute("aria-hidden", "true");
  frame.tabIndex = -1;
  frame.title = "print";
  // Laid out at a real size (a slide measures its words), out of sight.
  Object.assign(frame.style, {
    position: "fixed",
    left: "-12000px",
    top: "0",
    width: "1280px",
    height: "800px",
    border: "0",
    opacity: "0",
    pointerEvents: "none",
  });
  const origin = window.location.origin;
  return new Promise<ClientPrintOutcome>((resolve) => {
    let settled = false;
    const finish = (outcome: ClientPrintOutcome, keepFrame = false) => {
      if (settled) return;
      settled = true;
      window.clearTimeout(timer);
      window.removeEventListener("message", onMessage);
      if (!keepFrame) {
        frame.remove();
        if (active === frame) active = null;
      }
      resolve(outcome);
    };
    const onMessage = (event: MessageEvent) => {
      if (event.origin !== origin || event.source !== frame.contentWindow) return;
      const kind = (event.data as { type?: unknown } | null)?.type;
      if (kind === CLIENT_PRINT_WAITING) {
        frame.contentWindow?.postMessage({ type: CLIENT_PRINT_PAYLOAD, payload }, origin);
        return;
      }
      if (kind !== CLIENT_PRINT_DRAWN) return;
      const view = frame.contentWindow;
      if (!view) {
        finish("unavailable");
        return;
      }
      try {
        // The frame stays until its dialog closes (a browser that prints
        // asynchronously still reads it): `afterprint`, else a minute.
        const release = () => {
          frame.remove();
          if (active === frame) active = null;
        };
        view.addEventListener("afterprint", release, { once: true });
        window.setTimeout(release, 60_000);
        view.focus();
        view.print();
        finish("dialog", true);
      } catch {
        finish("unavailable");
      }
    };
    const timer = window.setTimeout(() => finish("unavailable"), options.timeoutMs ?? CLIENT_PRINT_DRAW_TIMEOUT_MS);
    window.addEventListener("message", onMessage);
    frame.src = `${localePrefix()}${CLIENT_PRINT_PATH}`;
    document.body.appendChild(frame);
  });
}
