/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { PDFName, type PDFDocument } from "pdf-lib";

export type PdfPageTabOrder = "R" | "C" | "S";

/** 1-based page number → declared `/Tabs` value; pages without one absent. */
export function readPageTabOrders(doc: PDFDocument): Map<number, PdfPageTabOrder> {
  const out = new Map<number, PdfPageTabOrder>();
  let pages: ReturnType<PDFDocument["getPages"]>;
  try {
    pages = doc.getPages();
  } catch {
    return out;
  }
  pages.forEach((page, index) => {
    try {
      const tabs = page.node.get(PDFName.of("Tabs"));
      if (!(tabs instanceof PDFName)) return;
      const value = tabs.decodeText();
      if (value === "R" || value === "C" || value === "S") {
        out.set(index + 1, value);
      }
    } catch {
      /* a malformed page keeps annotation order */
    }
  });
  return out;
}

/** Read the page tab orders from raw PDF bytes; empty for anything unreadable. */
export async function readPageTabOrdersFromBytes(
  bytes: ArrayBuffer | Uint8Array,
): Promise<Map<number, PdfPageTabOrder>> {
  try {
    const { loadPdfForInspection } = await import("./encrypted-guard");
    const doc = await loadPdfForInspection(bytes, { updateMetadata: false });
    if (!doc) return new Map();
    return readPageTabOrders(doc);
  } catch {
    return new Map();
  }
}

interface PlacedField {
  readonly percent: {
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
  };
}

/**
 * Order a page's widgets for its `/Tabs` value. Geometry is the rendered
 * percent box (top-left origin, already rotation-aware), so the order follows
 * what the reader sees. Rows / columns are formed by centre lines within half
 * the smaller widget's extent, so slightly misaligned boxes still share a
 * row. `/S`, an unknown value or an empty list keeps the given order.
 */
export function orderFieldsForTabs<T extends PlacedField>(
  fields: readonly T[],
  order: PdfPageTabOrder | null | undefined,
): T[] {
  if (fields.length < 2 || (order !== "R" && order !== "C")) return [...fields];
  const major = order === "R" ? "y" : "x";
  const minor = order === "R" ? "x" : "y";
  const majorExtent = order === "R" ? "height" : "width";
  const centre = (field: T, axis: "x" | "y") =>
    axis === "x"
      ? field.percent.x + field.percent.width / 2
      : field.percent.y + field.percent.height / 2;
  const byMajor = [...fields].sort(
    (a, b) => centre(a, major) - centre(b, major),
  );
  const groups: T[][] = [];
  let current: T[] = [];
  let currentCentre = 0;
  let currentExtent = 0;
  for (const field of byMajor) {
    const c = centre(field, major);
    const extent = Math.max(0, field.percent[majorExtent]);
    if (current.length === 0) {
      current = [field];
      currentCentre = c;
      currentExtent = extent;
      continue;
    }
    const tolerance = Math.max(0.25, Math.min(currentExtent, extent) / 2);
    if (Math.abs(c - currentCentre) <= tolerance) {
      current.push(field);
      continue;
    }
    groups.push(current);
    current = [field];
    currentCentre = c;
    currentExtent = extent;
  }
  if (current.length > 0) groups.push(current);
  return groups.flatMap((group) =>
    group.sort((a, b) => centre(a, minor) - centre(b, minor)),
  );
}
