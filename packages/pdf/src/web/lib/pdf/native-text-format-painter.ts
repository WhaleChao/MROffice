/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * 格式刷 for the PDF editor's text: copy one run's format and paint it onto the
 * next text clicked (WPS and Office semantics, as the Docs and presentation
 * painters keep them: one click paints once, a double click paints until
 * Escape or a click on the button again).
 *
 * Each line is edited by its own editor instance, which ends when the next
 * line opens, so the armed painter lives here, outside any one editor: the
 * editor that arms it writes it, the one that opens next reads it. A painter
 * belongs to one window (the page window's data-window-id); while it is
 * armed the window wears data-pdf-format-painter="armed" (the brush cursor)
 * and Escape anywhere disarms it before anything else sees the key.
 */

import { useSyncExternalStore } from "react";

import type { FormatPainterMode } from "@mrnote/core/web/lib/presentation-format-painter.ts";

/** What the painter carries, in units any run can take. */
export interface PdfNativeTextPaintFormat {
  /** The face's family name (matched by name in the target's fonts), or a
   * text box's portable face. Null keeps the target's face. */
  readonly fontFamily: string | null;
  readonly bold: boolean;
  readonly italic: boolean;
  /** The size on paper, in points. */
  readonly fontSizePoints: number | null;
  /** `#rrggbb`. */
  readonly color: string | null;
  readonly underline: boolean;
  readonly strikethrough: boolean;
  /** The superscript (positive) or subscript (negative) rise on paper. */
  readonly textRisePoints: number;
  /** A paragraph's own settings, painted onto a paragraph (WPS 格式刷 copies
   * them with the characters' format). */
  readonly paragraph?: Readonly<{
    alignment?: "left" | "center" | "right" | "justify" | "distribute";
    lineStep?: number;
    firstLineIndent?: number;
    leftIndent?: number;
    characterSpacingEm?: number;
  }>;
}

export interface PdfNativeTextFormatPainter {
  readonly windowId: string;
  readonly mode: FormatPainterMode;
  readonly format: PdfNativeTextPaintFormat;
}

let armed: PdfNativeTextFormatPainter | null = null;
const listeners = new Set<() => void>();

function windowElement(windowId: string): HTMLElement | null {
  if (typeof document === "undefined") return null;
  return Array.from(document.querySelectorAll<HTMLElement>("[data-window-id]"))
    .find((element) => element.dataset.windowId === windowId) ?? null;
}

// Escape cancels the painter before an editor or the reader sees the key
// (an open menu takes it first: it stops the key on its own capture).
function escapeDisarms(event: KeyboardEvent): void {
  if (event.key !== "Escape" || !armed) return;
  event.preventDefault();
  event.stopImmediatePropagation();
  disarmPdfNativeTextFormatPainter();
}

function publish(next: PdfNativeTextFormatPainter | null): void {
  const previous = armed;
  armed = next;
  if (previous && previous.windowId !== next?.windowId) {
    windowElement(previous.windowId)?.removeAttribute("data-pdf-format-painter");
  }
  if (next) windowElement(next.windowId)?.setAttribute("data-pdf-format-painter", "armed");
  if (typeof window !== "undefined") {
    if (next && !previous) window.addEventListener("keydown", escapeDisarms, true);
    if (!next && previous) window.removeEventListener("keydown", escapeDisarms, true);
  }
  for (const listener of listeners) listener();
}

export function armPdfNativeTextFormatPainter(painter: PdfNativeTextFormatPainter): void {
  publish(Object.freeze({ ...painter, format: Object.freeze({ ...painter.format }) }));
}

/** Disarm the painter (of `windowId` only, when one is named). */
export function disarmPdfNativeTextFormatPainter(windowId?: string): void {
  if (!armed || (windowId !== undefined && armed.windowId !== windowId)) return;
  publish(null);
}

/** The painter armed in this window, if any. */
export function readPdfNativeTextFormatPainter(
  windowId: string | null | undefined,
): PdfNativeTextFormatPainter | null {
  return windowId && armed?.windowId === windowId ? armed : null;
}

/** One paint done: a painter armed with one click is spent, a sticky one
 * stays armed. */
export function spendPdfNativeTextFormatPainter(windowId: string): void {
  if (armed?.windowId === windowId && armed.mode === "single") publish(null);
}

function subscribe(listener: () => void): () => void {
  listeners.add(listener);
  return () => listeners.delete(listener);
}

export function usePdfNativeTextFormatPainter(
  windowId: string | null | undefined,
): PdfNativeTextFormatPainter | null {
  return useSyncExternalStore(
    subscribe,
    () => readPdfNativeTextFormatPainter(windowId),
    () => null,
  );
}
