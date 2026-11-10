/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * 插入 › 智能图形 on a PDF page: the graphic the Docs editor inserts for the
 * same layout (the SmartArt engine's `document.insert`, its answer placed by
 * `documentSmartArtInsertAttrsFrom`, as DocumentToolbar places it), as a
 * managed object of the page (lib/pdf/managed-chart.ts, kind "smartart"),
 * its frame at the layout's own proportions.
 */
import { documentSmartArtInsertAttrsFrom } from "@mrnote/docs/web/components/console/editor/native-document/insertButtons.ts";
import type { OfficeSmartArtPlaceholders } from "@mrnote/core/web/lib/office-smartart/catalog.ts";
import { officeSmartArtSeedItemCount } from "@mrnote/core/web/lib/office-smartart/catalog.ts";
import { smartArtEngine } from "@mrnote/core/web/lib/office-smartart/remote.ts";
import type { PdfChart } from "./managed-chart";

/** PowerPoint's and Word's inserted SmartArt: at most this share of the page wide. */
const PDF_SMARTART_INSERT_WIDTH = 0.6;

export type PdfSmartArtInsertResult =
  | { readonly ok: true; readonly record: PdfChart }
  | { readonly ok: false; readonly reason: string; readonly memberCount?: number };

/** The frame (page shares) a graphic of `widthEmu` x `heightEmu` takes with its
 * top-left at (x, y) on a page `pageWidthPt` x `pageHeightPt` points: its own
 * proportions, as wide as the share allows and the page holds. */
export function pdfSmartArtInsertFrame(x: number, y: number, widthEmu: number, heightEmu: number,
  pageWidthPt: number, pageHeightPt: number): { x: number; y: number; width: number; height: number } {
  const left = Math.min(Math.max(0, x), 0.9);
  const top = Math.min(Math.max(0, y), 0.9);
  const aspect = widthEmu > 0 && heightEmu > 0 ? heightEmu / widthEmu : 0.6;
  let width = Math.min(PDF_SMARTART_INSERT_WIDTH, 1 - left);
  let height = width * pageWidthPt * aspect / pageHeightPt;
  if (top + height > 1) {
    height = 1 - top;
    width = height * pageHeightPt / aspect / pageWidthPt;
  }
  return { x: left, y: top, width, height };
}

export async function pdfSmartArtInsertRecord(input: {
  readonly layoutId: string;
  readonly placeholders: OfficeSmartArtPlaceholders;
  readonly page: number;
  readonly x: number;
  readonly y: number;
  readonly pageWidthPt: number;
  readonly pageHeightPt: number;
  readonly signal?: AbortSignal;
}): Promise<PdfSmartArtInsertResult> {
  // An alias card seeds its own sample (Vertical Bending Process: nine).
  const seedItems = officeSmartArtSeedItemCount(input.layoutId);
  if (seedItems === null) return { ok: false, reason: "unknown-layout" };
  const answer = await smartArtEngine("document.insert", {
    layoutId: input.layoutId,
    placeholders: {
      itemLabels: Array.from({ length: Math.max(1, seedItems) },
        (_, index) => input.placeholders.itemLabel(index + 1)),
      rootLabel: input.placeholders.rootLabel,
    },
  }, { signal: input.signal });
  const projection = documentSmartArtInsertAttrsFrom(answer.projection);
  // Never truncated into a different diagram: the caller says why.
  if (!projection.ok) return { ok: false, reason: projection.reason, memberCount: projection.memberCount };
  const attrs = projection.attrs as unknown as Record<string, unknown>;
  const frame = pdfSmartArtInsertFrame(input.x, input.y, Number(attrs.widthEmu), Number(attrs.heightEmu),
    input.pageWidthPt, input.pageHeightPt);
  return {
    ok: true,
    record: {
      kind: "smartart",
      id: `smartart-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 7)}`,
      page: input.page,
      ...frame,
      attrs,
      svg: null,
    },
  };
}
