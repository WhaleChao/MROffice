/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Plain-text projections of a deck: the text of one block, the Immersive
 * Reader projection of a slide, the bounded deck text handed to AI, and the
 * short block context/labels the AI surfaces and the animation pane show.
 *
 * Moved verbatim out of `components/notebook/contents/PresentationWindow.tsx`
 * (structural split, no behaviour change).
 */

import { compactAiContextText, evenlySpacedAiContextIndices } from "@mrnote/core/web/__cut__/mf740bff044.ts";
import { pictureObjectTokenSha, pictureObjectUrlSha } from "@mrnote/core/web/lib/page-picture-objects.ts";
import {
  normalizeBulletChar,
  normalizeNativeBulletReceipt,
  normalizeNumberFormat,
  renderParagraphs,
} from "@mrnote/core/web/lib/presentation-paragraphs.ts";
import { normalizeShapeKind } from "@mrnote/core/web/lib/presentation-shapes.ts";
import type {
  ImmersiveReaderProjection,
} from "@mrnote/docs/web/components/console/editor/native-document/DocumentImmersiveReader.tsx";
import { normalizeEmbedKind } from "@mrnote/core/web/lib/presentation-embed.ts";
import { geometryNumber } from "@mrnote/core/web/lib/presentation-object-buttons.ts";
import { normalizeParagraphsWithRuns } from "@mrnote/core/web/lib/presentation-text-buttons.ts";
import { normalizeShapeText } from "@mrnote/core/web/lib/presentation-text-format.ts";
import type {
  PresentationBlock,
  PresentationChartBlock,
} from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import type { SlideModel } from "@mrnote/core/web/lib/presentation/deck-model.ts";

/**
 * A chart's data as lines of text: its model's (`chartV2`, which is all an
 * imported or authored chart on the one model carries), else a pre-V2
 * block's own grid.
 */
function chartDataTextLines(block: PresentationChartBlock): string[] {
  const data = block.chartV2?.data;
  if (data?.kind === "category") {
    return [
      ["Category", ...data.series.map((series) => series.name)].join("\t"),
      ...data.categories.map((category, index) =>
        [category, ...data.series.map((series) => series.values[index] ?? "")].join("\t"),
      ),
    ];
  }
  if (data && "series" in data) {
    return data.series.map((series) => series.name).filter(Boolean);
  }
  return [
    ...(block.data?.headers?.length ? [block.data.headers.join("\t")] : []),
    ...(block.data?.rows || []).map((row) => row.join("\t")),
  ];
}

export function blockPlainText(block: PresentationBlock): string {
  if (block.type === "image") {
    return [block.alt, block.caption].filter(Boolean).join("\n").trim();
  }
  if (block.type === "chart") {
    const dataLines = chartDataTextLines(block);
    return [
      block.title,
      block.caption,
      block.chartType,
      ...(block.ranges || []),
      ...dataLines,
    ]
      .filter(Boolean)
      .join("\n")
      .trim();
  }
  if (block.type === "table") {
    return [
      block.caption,
      ...block.rows.map((row) => row.map((cell) => cell.text).join("\t")),
    ]
      .filter(Boolean)
      .join("\n")
      .trim();
  }
  if (block.type === "shape") {
    // A shape may carry a centered text label (PowerPoint "add text to shape");
    // surface it for AI context + plain-text search. A label-less shape is "".
    return normalizeShapeText(block.text) ?? "";
  }
  if (block.type === "embed") {
    return [block.title, block.url].filter(Boolean).join("\n").trim();
  }
  return block.text;
}

const PRESENTATION_IMMERSIVE_READER_MAX_CHARACTERS = 100_000;

/**
 * A block's words as the Immersive Reader reads them: a list's items led by
 * the marker the canvas draws for each (its number in the block's own
 * scheme, a bullet), as the shared reader leads a document's list items
 * (`DocumentImmersiveReader`). The plain extraction search and the AI use
 * carries no marker, so an automatic 1. / 2. list read as bare lines while
 * typed numbers stayed (PPT-B327); that extraction is left as it is.
 */
function presentationReaderBlockText(block: PresentationBlock): string {
  if (block.type !== "text" && block.type !== "shape") return blockPlainText(block);
  const paragraphs = normalizeParagraphsWithRuns(
    (block as { paragraphs?: unknown }).paragraphs,
  );
  if (!paragraphs?.length) return blockPlainText(block);
  const rendered = renderParagraphs(paragraphs, {
    bulletChar: normalizeBulletChar((block as { bulletChar?: string }).bulletChar),
    numberFormat: normalizeNumberFormat((block as { numberFormat?: string }).numberFormat),
    nativeBullet: normalizeNativeBulletReceipt(
      (block as { nativeBullet?: unknown }).nativeBullet,
      paragraphs.length,
    ),
  });
  if (!rendered.some((paragraph) => paragraph.marker)) return blockPlainText(block);
  return rendered
    .map((paragraph) => {
      if (!paragraph.marker || !paragraph.text.trim()) return paragraph.text;
      // A symbol-font glyph (a private-use code point) reads as a bullet.
      const marker = /[\uE000-\uF8FF]/u.test(paragraph.marker) ? "•" : paragraph.marker;
      return `${marker} ${paragraph.text}`;
    })
    .join("\n");
}

export function presentationImmersiveReaderProjection(
  slide: SlideModel,
): ImmersiveReaderProjection {
  const blocks: Array<ImmersiveReaderProjection["blocks"][number]> = [];
  let characters = 0;
  let unsupportedNodeCount = 0;
  let truncated = false;
  const append = (
    id: string,
    kind: ImmersiveReaderProjection["blocks"][number]["kind"],
    rawText: string,
    level?: number,
  ) => {
    const normalized = rawText.replace(/\u00a0/gu, " ").trim();
    if (!normalized) return;
    const available =
      PRESENTATION_IMMERSIVE_READER_MAX_CHARACTERS - characters;
    if (available <= 0) {
      truncated = true;
      return;
    }
    const text = normalized.slice(0, available);
    if (text.length < normalized.length) truncated = true;
    characters += text.length;
    blocks.push({ id, kind, text, ...(level ? { level } : {}) });
  };

  append(`presentation-slide-${slide.id}-title`, "heading", slide.title, 1);
  for (const block of slide.blocks) {
    const text = presentationReaderBlockText(block);
    if (!text.trim()) {
      unsupportedNodeCount += 1;
      continue;
    }
    append(
      `presentation-slide-${slide.id}-block-${block.id}`,
      block.type === "table" ? "table" : "paragraph",
      text,
    );
  }
  if (slide.notes?.trim()) {
    append(
      `presentation-slide-${slide.id}-notes`,
      "quote",
      slide.notes,
    );
  }
  return {
    blocks,
    plainText: blocks.map((block) => block.text).join("\n\n"),
    unsupportedNodeCount,
    truncated,
  };
}

// The deck text an AI turn sends, under the 200,000 characters the request
// carries (ai-panel/request-context.ts). At 18,000 / 80 slides a long deck
// reached the model as a few lines a slide.
const PRESENTATION_AI_DOCUMENT_MAX_CHARS = 180_000;

const PRESENTATION_AI_MAX_SAMPLED_SLIDES = 800;

const PRESENTATION_AI_MIN_SLIDE_CHARS = 140;

export function presentationDocumentTextForAi(
  slides: readonly SlideModel[],
  activeSlideIndex: number,
): { text: string; truncated: boolean; sampledSlideCount: number } {
  const normalizedActiveIndex =
    activeSlideIndex >= 0 && activeSlideIndex < slides.length
      ? activeSlideIndex
      : null;
  const indices = evenlySpacedAiContextIndices(
    slides.length,
    Math.min(PRESENTATION_AI_MAX_SAMPLED_SLIDES, slides.length),
    normalizedActiveIndex,
  );
  const omittedSlideCount = slides.length - indices.length;
  const header = [
    `[Presentation · slides=${slides.length}]`,
    normalizedActiveIndex === null
      ? ""
      : `[Active slide ${normalizedActiveIndex + 1}]`,
    omittedSlideCount > 0
      ? `[… ${omittedSlideCount} slides omitted by even deck sampling …]`
      : "",
  ]
    .filter(Boolean)
    .join("\n");
  const fullSections = indices.map((index) => {
    const slide = slides[index]!;
    const body = slide.blocks.map(blockPlainText).filter(Boolean).join("\n");
    const notes = String(slide.notes || "").trim();
    return [
      `[Slide ${index + 1}: ${slide.title || "Untitled"}${index === normalizedActiveIndex ? " · active" : ""}]`,
      body,
      notes ? `[Speaker notes]\n${notes}` : "",
    ]
      .filter(Boolean)
      .join("\n");
  });
  const separators = Math.max(0, fullSections.length - 1) * 2;
  const perSlideChars = Math.max(
    PRESENTATION_AI_MIN_SLIDE_CHARS,
    Math.floor(
      (PRESENTATION_AI_DOCUMENT_MAX_CHARS - header.length - separators - 2) /
        Math.max(1, fullSections.length),
    ),
  );
  const sections = fullSections.map((section) =>
    compactAiContextText(
      section,
      perSlideChars,
      "\n…[slide middle omitted]…\n",
    ),
  );
  const joined = [header, ...sections].filter(Boolean).join("\n\n");
  return {
    text: compactAiContextText(
      joined,
      PRESENTATION_AI_DOCUMENT_MAX_CHARS,
      "\n…[deck middle omitted]…\n",
    ),
    truncated:
      omittedSlideCount > 0 ||
      fullSections.some((section) => section.length > perSlideChars) ||
      joined.length > PRESENTATION_AI_DOCUMENT_MAX_CHARS,
    sampledSlideCount: indices.length,
  };
}

function shortContextString(
  value: unknown,
  maxLength: number,
): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  if (!trimmed) return undefined;
  return trimmed.slice(0, maxLength);
}

function presentationImageSourceContext(src: string): Record<string, unknown> {
  if (pictureObjectUrlSha(src) !== null || pictureObjectTokenSha(src) !== null) {
    // A picture stored beside the deck (lib/page-picture-objects).
    return { image_src_kind: "embedded", image_src_prefix: "stored-picture" };
  }
  if (src.startsWith("data:image/")) {
    const headerEnd = src.indexOf(",");
    return {
      image_src_kind: "embedded",
      image_src_prefix:
        headerEnd > 0 ? src.slice(0, headerEnd + 1) : "data:image/",
    };
  }
  if (src.startsWith("https://") || src.startsWith("http://")) {
    return {
      image_src_kind: "remote",
      image_src: src.slice(0, 520),
    };
  }
  return { image_src_kind: src ? "other" : "none" };
}

export function presentationBlockContext(
  block: PresentationBlock,
): Record<string, unknown> {
  const context: Record<string, unknown> = {
    block_id: block.id,
    block_type: block.type,
  };
  const left = geometryNumber(block, "left");
  const top = geometryNumber(block, "top");
  const width = geometryNumber(block, "width");
  const height = geometryNumber(block, "height");
  const rotation = geometryNumber(block, "rotation");
  const sourceOrder = geometryNumber(block, "sourceOrder");
  const zIndex = geometryNumber(block, "zIndex");
  if (left !== null) context.block_left = left;
  if (top !== null) context.block_top = top;
  if (width !== null) context.block_width = width;
  if (height !== null) context.block_height = height;
  if (rotation !== null) context.block_rotation = rotation;
  if (sourceOrder !== null) context.block_source_order = sourceOrder;
  if (zIndex !== null) context.block_z_index = zIndex;
  const source = (block as Record<string, unknown>).source;
  if (source && typeof source === "object") {
    context.source = source;
  }
  if (block.type === "image") {
    const alt = shortContextString(block.alt, 500);
    const caption = shortContextString(block.caption, 500);
    if (alt) context.image_alt = alt;
    if (caption) context.image_caption = caption;
    Object.assign(context, presentationImageSourceContext(block.src || ""));
  } else if (block.type === "chart") {
    const title = shortContextString(block.title, 500);
    const caption = shortContextString(block.caption, 500);
    const chartType = shortContextString(block.chartType, 80);
    if (title) context.chart_title = title;
    if (caption) context.chart_caption = caption;
    if (chartType) context.chart_type = chartType;
    if (Array.isArray(block.ranges) && block.ranges.length > 0) {
      context.chart_ranges = block.ranges.slice(0, 10);
      context.chart_range_count = block.ranges.length;
      context.chart_ranges_truncated = block.ranges.length > 10;
    }
    // The grid's shape: the model's (a category per row, a column per
    // series and one for the categories), else a pre-V2 block's own grid.
    const modelData = block.chartV2?.data;
    context.chart_row_count =
      modelData?.kind === "category" ? modelData.categories.length : block.data?.rows?.length || 0;
    context.chart_column_count =
      modelData?.kind === "category" ? modelData.series.length + 1 : block.data?.headers?.length || 0;
  } else if (block.type === "table") {
    const caption = shortContextString(block.caption, 500);
    if (caption) context.table_caption = caption;
    context.table_row_count = block.rows.length;
    context.table_column_count = Math.max(
      0,
      ...block.rows.map((row) => row.length),
    );
  } else if (block.type === "shape") {
    context.shape_kind = normalizeShapeKind(block.shape);
    if (typeof block.fill === "string") context.shape_fill = block.fill;
    if (typeof block.stroke === "string") context.shape_stroke = block.stroke;
  } else if (block.type === "embed") {
    context.embed_kind = normalizeEmbedKind(block.kind);
    const title = shortContextString(block.title, 500);
    if (title) context.embed_title = title;
    const url = shortContextString(block.url, 520);
    if (url) context.embed_url = url;
  } else {
    const preview = shortContextString(block.text, 500);
    if (preview) context.text_preview = preview;
  }
  return context;
}

/**
 * A block's short human label: its words, else its title, else the name the
 * author gave it in the Selection Pane, else `fallbackName` (the editor hands
 * in the block's accessible name, e.g. "Shape block: Rectangle"), and only
 * then its internal id. A text-less shape thus reads as the Selection Pane
 * names it instead of as `block-…` (PPT audit B442); the id is kept as the
 * stable target, this label is display only.
 */
export function presentationBlockLabel(
  block: PresentationBlock,
  fallbackName?: string,
): string {
  return (
    shortContextString(blockPlainText(block), 120) ||
    shortContextString((block as Record<string, unknown>).title, 120) ||
    shortContextString((block as Record<string, unknown>).name, 120) ||
    shortContextString(fallbackName, 120) ||
    block.id
  );
}
