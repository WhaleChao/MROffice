/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What a slide block says about itself (its placed geometry, whether it came
 * from an imported PPTX, whether it is a placeholder), apart from the slide
 * commands that use them (presentation-object-buttons and
 * presentation-shape-buttons re-export these): the text layout reads them for
 * a spreadsheet's shapes too, and the command modules brought the deck's
 * layouts, clipboard, copy and slide commands with them.
 */
import type { PresentationBlock } from "../components/console/editor/native-document/documentCommands";
import { asRecord } from "./as-record";

export function geometryNumber(block: PresentationBlock, key: string): number | null {
  const value = (block as Record<string, unknown>)[key];
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

export function hasLayoutGeometry(block: PresentationBlock): boolean {
  return (
    geometryNumber(block, "left") !== null ||
    geometryNumber(block, "top") !== null
  );
}

export function isImportedPptxBlock(block: PresentationBlock): boolean {
  const record = block as Record<string, unknown>;
  const source = asRecord(record.source);
  const hasNativeSource =
    (typeof source?.slideId === "string" && source.slideId.trim() !== "") ||
    // XML-only ChartEx imports retain the slide part instead of a slide ID.
    // Their explicit native custody must not look like a scratch chart.
    (block.type === "chart" && source?.nativeChartDisposition === "preserve-only" &&
      typeof source.slidePart === "string" && /^ppt\/slides\/[^/]+\.xml$/u.test(source.slidePart)) ||
    source?.inheritedFrom === "layout" ||
    source?.inheritedFrom === "master";
  return (
    hasNativeSource &&
    typeof record.sourceOrder === "number" &&
    Number.isInteger(record.sourceOrder) &&
    typeof record.zIndex === "number" &&
    record.sourceOrder === record.zIndex
  );
}

/**
 * Whether a block is a layout placeholder, read off ANY block kind: the
 * importer projects a text-bearing placeholder whose label it admits as a
 * SHAPE block stamped with `placeholder` / `placeholderType` / `role`, and
 * Change Shape produces the same from a scratch title, so the rule asks the
 * block, not its type. A placeholder is never grouped: Microsoft states it
 * (a shape cannot be grouped with the title or content placeholder) and
 * PowerPoint writes `a:spLocks noGrp="1"` on every one.
 */
export function presentationBlockIsPlaceholder(block: PresentationBlock): boolean {
  const record = block as PresentationBlock & {
    placeholder?: unknown;
    placeholderType?: unknown;
    role?: unknown;
  };
  return [record.placeholder, record.placeholderType, record.role].some(
    (value) => typeof value === "string" && value.trim().length > 0,
  );
}

export function isImportedPptxTextBlock(block: PresentationBlock): boolean {
  if (block.type !== "text") return false;
  const record = block as Record<string, unknown>;
  // Imported text carries its native custody (the slide and shape it came
  // from, or the layout / master it inherits from) beside its paint order.
  // The AI composer writes `sourceOrder` = `zIndex` on every block it makes,
  // and on paint order alone an AI deck's text read as imported: drawn and
  // edited through the imported overlay, kept out of the Outline View as a
  
  const source =
    record.source && typeof record.source === "object" && !Array.isArray(record.source)
      ? (record.source as Record<string, unknown>)
      : null;
  const hasNativeSource =
    (typeof source?.slideId === "string" && source.slideId.trim() !== "") ||
    source?.inheritedFrom === "layout" ||
    source?.inheritedFrom === "master";
  return (
    hasNativeSource &&
    typeof record.sourceOrder === "number" &&
    Number.isInteger(record.sourceOrder) &&
    typeof record.zIndex === "number" &&
    record.sourceOrder === record.zIndex
  );
}
