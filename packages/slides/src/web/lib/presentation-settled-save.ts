/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What the server settled on an imported deck's page after a save, adopted
 * by a REST editor field by field.
 *
 * A save that writes an Office theme over the file's master settles the page,
 * in the same version, to the form the written file reopens in
 * (`pptx_theme_settle.settle_page_after_theme_write`): each listed slide's
 * master graphics are the ones the written master draws, and the theme's
 * identity is the package's (`origin: "package"`); a slide a layout change
 * rebound takes its new layout's graphics (and the master's it shows). The save receipt names
 * them (`editable_page_patchback.settled`). The editor takes exactly those
 * fields from the saved body into the draft it sent -- nothing else of the
 * saved body is read, nothing is inferred -- and only for the version the
 * receipt names.
 */

import { pageDocumentJsonEquals } from "@mrnote/core/web/lib/page-save-acknowledgement.ts";
import { PRESENTATION_DECK_RECEIPTS_KEY } from "@mrnote/core/web/lib/presentation-deck-receipts.ts";

export const PRESENTATION_SETTLE_SCHEMA_VERSION = "pptx-page-settle-v1";

type Rec = Record<string, unknown>;

const record = (value: unknown): Rec | undefined =>
  value && typeof value === "object" && !Array.isArray(value) ? (value as Rec) : undefined;

type SettledKind = "masterDecorations" | "layoutDecorations" | "bandPlaceholders" | "themePainted";
const SETTLED_KINDS: SettledKind[] = ["masterDecorations", "layoutDecorations", "bandPlaceholders", "themePainted"];

export interface PresentationSettledReceipt {
  contentVersion: number;
  /** Per slide, its master graphics, its layout graphics and/or its footer
   * band placeholders (Header & Footer on a slide holding its band as
   * blocks) now, in order; and the blocks whose colours and fonts follow the
   * theme just picked, as the written file reads them (``themePainted``). */
  slides: ({ slideId: string } & Partial<Record<SettledKind, string[]>>)[];
  themeOrigin?: "package";
}

/** The receipt as the server states it, or null when it is not one. */
export function readPresentationSettledReceipt(value: unknown): PresentationSettledReceipt | null {
  const receipt = record(value);
  if (!receipt || receipt.schemaVersion !== PRESENTATION_SETTLE_SCHEMA_VERSION) return null;
  const { contentVersion, slides, themeOrigin } = receipt;
  if (typeof contentVersion !== "number" || !Number.isSafeInteger(contentVersion) || contentVersion < 1) return null;
  if (!Array.isArray(slides) || (themeOrigin !== undefined && themeOrigin !== "package")) return null;
  const read: PresentationSettledReceipt["slides"] = [];
  const seen = new Set<string>();
  for (const raw of slides) {
    const slide = record(raw);
    if (!slide || typeof slide.slideId !== "string" || !slide.slideId || seen.has(slide.slideId)) return null;
    const entry: PresentationSettledReceipt["slides"][number] = { slideId: slide.slideId };
    for (const kind of SETTLED_KINDS) {
      const ids = slide[kind];
      if (ids === undefined) continue;
      if (!Array.isArray(ids) || ids.some((id) => typeof id !== "string" || !id) || new Set(ids).size !== ids.length) {
        return null;
      }
      entry[kind] = [...(ids as string[])];
    }
    if (!entry.masterDecorations && !entry.layoutDecorations && !entry.bandPlaceholders && !entry.themePainted) return null;
    seen.add(slide.slideId);
    read.push(entry);
  }
  if (!read.length && themeOrigin === undefined) return null;
  return { contentVersion, slides: read, ...(themeOrigin ? { themeOrigin } : {}) };
}

/** A master's (layout's) own graphic drawn on a slide (not one of its placeholders). */
function isGraphic(block: unknown, origin: "master" | "layout"): boolean {
  const value = record(block);
  const source = record(value?.source);
  return !!value && source?.inheritedFrom === origin && !value.placeholderType && !source.placeholderType;
}
const isMasterGraphic = (block: unknown) => isGraphic(block, "master");

/** A footer band placeholder the slide holds as its own block. */
function isBandPlaceholder(block: unknown): boolean {
  const value = record(block);
  const source = record(value?.source);
  const kind = value?.placeholderType ?? source?.placeholderType;
  return !!value && !source?.inheritedFrom && (kind === "ftr" || kind === "sldNum" || kind === "dt");
}

function slidesById(content: Rec): Map<string, Rec> | null {
  if (!Array.isArray(content.slides)) return null;
  const map = new Map<string, Rec>();
  for (const raw of content.slides) {
    const slide = record(raw);
    if (!slide || typeof slide.id !== "string" || map.has(slide.id) || !Array.isArray(slide.blocks)) return null;
    map.set(slide.id, slide);
  }
  return map;
}

/**
 * ``draft`` (the stored-form body the save sent) with the settled fields the
 * receipt names taken from ``saved`` (the body the server stored for
 * ``savedVersion``): each listed slide's master graphics, in their place
 * behind the slide, and the theme's package origin. Null when the receipt is
 * not for that version, names a slide or a graphic either body lacks, or the
 * theme it settles is not the draft's -- the draft is then left alone, and
 * the next save settles it again.
 */
export function readPresentationSettledSave(
  draft: Rec,
  saved: unknown,
  savedVersion: number,
  receiptValue: unknown,
): Rec | null {
  const receipt = readPresentationSettledReceipt(receiptValue);
  const savedContent = record(saved);
  if (!receipt || !savedContent || receipt.contentVersion !== savedVersion) return null;
  const draftSlides = slidesById(draft), savedSlides = slidesById(savedContent);
  if (!draftSlides || !savedSlides) return null;
  const next: Rec = { ...draft, slides: [...(draft.slides as unknown[])] };
  for (const settled of receipt.slides) {
    const slide = draftSlides.get(settled.slideId), stored = savedSlides.get(settled.slideId);
    if (!slide || !stored) return null;
    const storedBlocks = new Map((stored.blocks as unknown[]).map((block) => [record(block)?.id, block]));
    let blocks = slide.blocks as unknown[];
    for (const kind of SETTLED_KINDS) {
      const ids = settled[kind];
      if (!ids) continue;
      if (kind === "themePainted") {
        // These blocks as the server stored them, each where it stands.
        const listed = new Set(ids);
        if (ids.some((id) => !storedBlocks.has(id))) return null;
        blocks = blocks.map((block) => {
          const id = record(block)?.id;
          return typeof id === "string" && listed.has(id) ? structuredClone(storedBlocks.get(id)) : block;
        });
        continue;
      }
      if (kind === "bandPlaceholders") {
        // The band the written slide has: a placeholder the page holds keeps
        // its place, one the writer added (the slide number) comes last.
        const band = ids.map((id) => storedBlocks.get(id));
        if (band.some((block) => !isBandPlaceholder(block))) return null;
        const listed = new Map(ids.map((id, index) => [id, band[index]]));
        const placed = new Set<string>();
        const kept: unknown[] = [];
        for (const block of blocks) {
          if (!isBandPlaceholder(block)) { kept.push(block); continue; }
          const id = record(block)?.id as string;
          if (listed.has(id)) { kept.push(structuredClone(listed.get(id))); placed.add(id); }
        }
        for (const [id, block] of listed) if (!placed.has(id)) kept.push(structuredClone(block));
        blocks = kept;
        continue;
      }
      const origin = kind === "masterDecorations" ? "master" : "layout";
      const graphics = ids.map((id) => storedBlocks.get(id));
      if (graphics.some((block) => !isGraphic(block, origin))) return null;
      // The master draws first, then the layout, behind the slide's own
      // shapes -- where the server put them.
      const first = blocks.findIndex((block) => isGraphic(block, origin));
      const fallback = origin === "master" ? 0 : Math.max(0, blocks.findIndex((block) => !isMasterGraphic(block)));
      const at = first >= 0 ? first : fallback;
      const kept = blocks.filter((block) => !isGraphic(block, origin));
      kept.splice(Math.min(at, kept.length), 0, ...graphics.map((block) => structuredClone(block)));
      blocks = kept;
    }
    const index = (next.slides as unknown[]).indexOf(slide);
    (next.slides as unknown[])[index] = { ...slide, blocks };
  }
  if (receipt.themeOrigin) {
    const theme = record(draft.theme), identity = record(theme?.identity);
    const savedIdentity = record(record(savedContent.theme)?.identity);
    if (!theme || !identity || !savedIdentity || savedIdentity.origin !== "package" ||
      !pageDocumentJsonEquals({ ...identity, origin: "package" }, savedIdentity)) return null;
    next.theme = { ...theme, identity: { ...identity, origin: "package" } };
  }
  // The graphics may name a stored picture or drawing by its table entry.
  for (const key of ["media", PRESENTATION_DECK_RECEIPTS_KEY]) {
    const table = record(savedContent[key]);
    if (table) next[key] = { ...(record(draft[key]) ?? {}), ...table };
  }
  return next;
}
