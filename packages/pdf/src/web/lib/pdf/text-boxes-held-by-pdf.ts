/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfTextBoxEdit } from "../../components/notebook/contents/pdf/PdfTextBoxOverlay";

import { pdfPageContentSchema } from "@mrnote/core/web/lib/native-page-schemas.ts";

import { isNativePdfFreeTextRecord } from "./native-free-text-annotations";
import { isManagedNativeTextBoxEditEligible, parseManagedNativeTextBoxShadow } from "./native-text-box-roundtrip";
import { getTextBoxEdits } from "./page-content-model";

/** A managed text box's row is named after its box
 * (planManagedNativeTextBoxAnnotations). */
const ROW_OF_BOX = "__mrnote_native_text_box__";

/** A box the reader made whose place in the PDF is a row of its own: one
 * with something to draw, that the native writer takes. (Not a line of the
 * PDF's own text turned into a box, nor a text box the PDF came with: those
 * are written through their own rows.) */
function waitsForItsOwnRow(row: PdfTextBoxEdit): boolean {
  if (!row.created || row.deleted || row.sourceText || isNativePdfFreeTextRecord(row)) return false;
  const drawn = row.text.trim().length > 0 || (row.kind === "formula" && Boolean(row.latex?.trim()));
  return drawn && isManagedNativeTextBoxEditEligible(row);
}

type Reading = Readonly<{ rows: PdfTextBoxEdit[]; unwritten: ReadonlySet<string> }>;

const readings = new WeakMap<object, Reading>();
const NOTHING_UNWRITTEN: ReadonlySet<string> = new Set();

function read(content: Record<string, unknown> | null | undefined): Reading {
  const cached = content ? readings.get(content) : undefined;
  if (cached) return cached;
  const all = getTextBoxEdits(content ?? undefined);
  let reading: Reading = { rows: all, unwritten: NOTHING_UNWRITTEN };
  // A record that does not read as a PDF page's says nothing about its
  // rows: every box of its list stands as saved, as before.
  const parsed = all.some(waitsForItsOwnRow) ? pdfPageContentSchema.safeParse(content) : null;
  if (parsed?.success) {
    const written = new Set<string>();
    for (const annotation of parsed.data.annotations as unknown[]) {
      const id = (annotation as { id?: unknown } | null)?.id;
      if (typeof id !== "string" || !id.startsWith(ROW_OF_BOX)) continue;
      // The box's row, whatever version of the directive it is written in
      // (one this build does not read is still the box's place in the PDF);
      // a row that takes the box out of the PDF is not.
      const record = parseManagedNativeTextBoxShadow(annotation);
      if (record && "deleted" in record && record.deleted) continue;
      written.add(id.slice(ROW_OF_BOX.length));
    }
    const unwritten = new Set<string>();
    const rows = all.filter((row) => {
      if (!waitsForItsOwnRow(row) || written.has(row.id)) return true;
      unwritten.add(`${row.page}:${row.id}`);
      return false;
    });
    if (unwritten.size > 0) reading = { rows, unwritten };
  }
  if (content) readings.set(content, reading);
  return reading;
}

/**
 * The record's text boxes without those the reader made that the PDF holds
 * no row of: what the saved page holds of the page's boxes, for everything
 * that asks whether a box is saved.
 */
export function pdfTextBoxRowsHeldByPdf(content: Record<string, unknown> | null | undefined): PdfTextBoxEdit[] {
  return read(content).rows;
}

/**
 * The boxes left out by pdfTextBoxRowsHeldByPdf, by `${page}:${id}`: in the
 * record's list, not in the PDF.
 */
export function pdfTextBoxRowsOnlyInRecord(content: Record<string, unknown> | null | undefined): ReadonlySet<string> {
  return read(content).unwritten;
}
