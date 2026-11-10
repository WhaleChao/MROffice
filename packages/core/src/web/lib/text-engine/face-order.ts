/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Which of two FontFace-API faces of one family wins a character both cover.
 *
 * CSS Fonts 4 §4.5 has the last-defined @font-face rule checked first; faces
 * added through the FontFace API do not follow that in every engine (the
 * registry's Noto Sans SC slices overlap -- a frequency slice and the Latin
 * slice both hold "«", from files of different versions -- and the browser
 * drew the earlier-added one). Rather than assume, the page asks once: two
 * served faces with one overlapping character, measured on a canvas.
 */
import { faceMetrics } from "./face-metrics";

export type ApiFaceOrder = "last-wins" | "first-wins";

const PROBE_FAMILY = "mrai-text-engine-order-probe";
const FIRST = "/fonts/office/carlito/400-latin.woff2";
const SECOND = "/fonts/office/arimo/400-latin.woff2";

let order: ApiFaceOrder | "pending" | undefined;
const listeners = new Set<() => void>();

export function onApiFaceOrderKnown(listener: () => void): () => void {
  listeners.add(listener);
  return () => listeners.delete(listener);
}

/** Tests: fix the order. */
export function setApiFaceOrderForTests(value: ApiFaceOrder | undefined): void {
  order = value;
}

/** The order, or "pending" while the page is asked. */
export function apiFaceOrder(): ApiFaceOrder | "pending" {
  if (order) return order;
  if (typeof document === "undefined" || !document.fonts || typeof FontFace === "undefined") return (order = "last-wins");
  order = "pending";
  void probe().then((answer) => {
    order = answer;
    for (const listener of [...listeners]) {
      try { listener(); } catch { /* its own */ }
    }
  });
  return order;
}

async function probe(): Promise<ApiFaceOrder> {
  const base = typeof document.baseURI === "string" ? document.baseURI : undefined;
  const first = new FontFace(PROBE_FAMILY, `url(${FIRST}) format("woff2")`, { unicodeRange: "U+41" });
  const second = new FontFace(PROBE_FAMILY, `url(${SECOND}) format("woff2")`, { unicodeRange: "U+41" });
  try {
    faceMetrics(new URL(FIRST, base).href);
    faceMetrics(new URL(SECOND, base).href);
    await Promise.all([first.load(), second.load()]);
    document.fonts.add(first);
    document.fonts.add(second);
    const context = document.createElement("canvas").getContext("2d");
    if (!context) return "last-wins";
    context.font = `400 1000px "${PROBE_FAMILY}"`;
    const width = context.measureText("A").width;
    // Carlito's A is 0.579 em, Arimo's 0.667 em.
    return Math.abs(width - 667) < Math.abs(width - 579) ? "last-wins" : "first-wins";
  } catch {
    return "last-wins";
  } finally {
    document.fonts.delete(first);
    document.fonts.delete(second);
  }
}
