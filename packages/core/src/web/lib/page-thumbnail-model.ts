/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What a page's thumbnail draws, read from the page's own content: the first
 * blocks of a document, the top-left of a workbook's first sheet, the first
 * lines of a code page. Pure and bounded, so a large page costs a walk of
 * its first screenful, not of the whole document.
 */

import {
  normalizePageOrientation,
  pageGeometryTwips,
} from "../components/console/editor/native-document/pageSetup";
import {
  emuToPx,
  normalizeFloatAnchor,
  placeFloatOnPage,
  type FloatPageGeometry,
} from "./office-float-anchor";

export type DocumentThumbnailBlock =
  | { tag: "h1" | "h2" | "h3" | "p" | "li" | "oli" | "task" | "quote" | "code"; text: string; checked?: boolean }
  | { tag: "hr" }
  | { tag: "img"; src: string }
  | { tag: "table"; rows: string[][] }
  /** A drawing in the text's flow (an in-line shape): its own box, in px. */
  | { tag: "shape"; attrs: Record<string, unknown>; widthPx: number; heightPx: number };

/**
 * A drawing Word places on the page rather than in the text (`wp:anchor`):
 * where it lands on the first page, in CSS px from the page's top-left, as
 * the editor's own float placement puts it (`placeFloatOnPage`).
 */
export interface DocumentThumbnailFloat {
  kind: "shape" | "image";
  /** The shape's attrs (`documentShape`), drawn by the editor's painter. */
  attrs: Record<string, unknown>;
  /** The picture's address (an image). */
  src?: string;
  leftPx: number;
  topPx: number;
  widthPx: number;
  heightPx: number;
  /** Behind Text: under the words. */
  behind: boolean;
  /** `@relativeHeight`: a larger one is drawn later. */
  z: number;
}

/** A document's first page: its size, its flowing blocks and its drawings. */
export interface DocumentThumbnailPage {
  widthPx: number;
  heightPx: number;
  blocks: DocumentThumbnailBlock[];
  floats: DocumentThumbnailFloat[];
}

export interface DocumentThumbnailOptions {
  /** Flowing blocks drawn at most. */
  limit?: number;
  /**
   * The address a picture is drawn from: an imported Word document names its
   * pictures by digest (`mrai-picture:<sha256>`), drawn from the page's own
   * picture route (`documentPictureDisplaySrc`). Null: not drawn.
   */
  pictureSrc?: (src: string) => string | null;
}

const MAX_TEXT = 320;
const MAX_TABLE_ROWS = 6;
const MAX_TABLE_COLS = 5;
const MAX_FLOATS = 80;
const SHAPE_TYPES = new Set(["documentShape", "inlineDocumentShape"]);
const IMAGE_TYPES = new Set(["image", "imageBlock", "inlineImage", "figure"]);

function record(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

function children(node: Record<string, unknown>): Record<string, unknown>[] {
  return Array.isArray(node.content)
    ? (node.content.map(record).filter(Boolean) as Record<string, unknown>[])
    : [];
}

/** The node's text, inline breaks as spaces, cut to a thumbnail's worth. */
function textOf(node: Record<string, unknown>, budget = MAX_TEXT): string {
  let out = "";
  const walk = (current: Record<string, unknown>) => {
    if (out.length >= budget) return;
    if (current.type === "text" && typeof current.text === "string") {
      out += current.text;
      return;
    }
    if (current.type === "hardBreak") {
      out += " ";
      return;
    }
    for (const child of children(current)) walk(child);
  };
  walk(node);
  return out.replace(/\s+/g, " ").trim().slice(0, budget);
}

/** A thumbnail can show an image the page can: an absolute or same-site URL. */
function imageSource(value: unknown): string | null {
  if (typeof value !== "string" || !value) return null;
  if (/^(https?:|data:image\/|blob:|\/)/i.test(value)) return value;
  return null;
}

function finitePositive(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) && value > 0 ? value : null;
}

const EMU_PER_TWIP = 635;

function pageGeometry(attrs: Record<string, unknown>): FloatPageGeometry {
  const twips = pageGeometryTwips(
    attrs.pageSize,
    attrs.pageMargins,
    attrs.customPageSize,
    attrs.customPageMargins,
    normalizePageOrientation(attrs.pageOrientation),
  );
  const px = (value: number) => emuToPx(value * EMU_PER_TWIP);
  const mode = attrs.wordCompatibilityMode;
  return {
    pageWidthPx: px(twips.width),
    pageHeightPx: px(twips.height),
    marginLeftPx: px(twips.left),
    marginTopPx: px(twips.top),
    marginRightPx: px(twips.right),
    marginBottomPx: px(twips.bottom),
    keepOnPage: {
      wordCompatibilityMode: typeof mode === "number" && Number.isFinite(mode) ? mode : null,
      headerFooter: false,
    },
  };
}

/**
 * A document's first page (ProseMirror JSON), in reading order: its blocks
 * up to the end of the page -- a page break, a section break that starts a
 * new page, or a block that starts its own page ends it, as the page itself
 * ends there (a cover page is drawn as the cover, not run on into the
 * chapters after it) -- and the drawings Word places on it (a designed cover
 * made of shapes is all drawings, no text at all).
 */
export function documentThumbnailPage(
  content: unknown,
  options: DocumentThumbnailOptions = {},
): DocumentThumbnailPage {
  const limit = options.limit ?? 28;
  const root = record(content);
  const geometry = pageGeometry(record(root?.attrs) ?? {});
  const blocks: DocumentThumbnailBlock[] = [];
  const floats: DocumentThumbnailFloat[] = [];
  let pageEnded = false;
  const picture = (value: unknown): string | null => {
    if (typeof value !== "string" || !value) return null;
    return options.pictureSrc?.(value) ?? imageSource(value);
  };
  const push = (block: DocumentThumbnailBlock) => {
    if (blocks.length < limit) blocks.push(block);
  };
  const shown = () => floats.length > 0 || blocks.some((block) => block.tag !== "p" || block.text);
  /** A drawing placed on the page: kept as a float. */
  const placed = (node: Record<string, unknown>): boolean => {
    const type = String(node.type ?? "");
    const attrs = record(node.attrs) ?? {};
    const anchor = normalizeFloatAnchor(attrs.floatAnchor);
    if (!anchor) return false;
    if (anchor.hidden || attrs.hidden === true || floats.length >= MAX_FLOATS) return true;
    const src = IMAGE_TYPES.has(type) ? picture(attrs.src) : null;
    if (!SHAPE_TYPES.has(type) && !src) return true;
    const place = placeFloatOnPage(anchor, geometry);
    // A shape's picture fill is named like any picture of the page.
    const fill = record(attrs.fillPicture);
    const fillSrc = fill && typeof fill.src === "string" ? options.pictureSrc?.(fill.src) : null;
    floats.push({
      kind: src ? "image" : "shape",
      attrs: fillSrc ? { ...attrs, fillPicture: { ...fill, src: fillSrc } } : attrs,
      ...(src ? { src } : {}),
      leftPx: place.leftPx,
      topPx: place.topPx,
      widthPx: place.widthPx,
      heightPx: place.heightPx,
      behind: anchor.behindDoc,
      z: anchor.relativeHeight,
    });
    return true;
  };
  /** A drawing in the flow: an in-line picture or shape. */
  const inline = (node: Record<string, unknown>): boolean => {
    const type = String(node.type ?? "");
    const attrs = record(node.attrs) ?? {};
    if (SHAPE_TYPES.has(type)) {
      const width = finitePositive(attrs.widthEmu);
      const height = finitePositive(attrs.heightEmu);
      if (width && height && attrs.hidden !== true) {
        push({ tag: "shape", attrs, widthPx: emuToPx(width), heightPx: emuToPx(height) });
      }
      return true;
    }
    if (IMAGE_TYPES.has(type)) {
      const src = picture(attrs.src);
      if (src) {
        push({ tag: "img", src });
        return true;
      }
    }
    return false;
  };
  /** The drawings a paragraph holds among its words. */
  const drawingsIn = (node: Record<string, unknown>): Record<string, unknown>[] => {
    const found: Record<string, unknown>[] = [];
    const visit = (current: Record<string, unknown>, depth: number) => {
      if (depth > 4 || found.length >= 12) return;
      for (const child of children(current)) {
        const type = String(child.type ?? "");
        if (SHAPE_TYPES.has(type) || IMAGE_TYPES.has(type)) found.push(child);
        else visit(child, depth + 1);
      }
    };
    visit(node, 0);
    return found;
  };
  const walk = (node: Record<string, unknown>, listKind: "li" | "oli" | null) => {
    if (pageEnded || blocks.length >= limit) return;
    const type = String(node.type ?? "");
    const attrs = record(node.attrs) ?? {};
    // a page that shows nothing yet (blank lines before a break) is not the
    // page the thumbnail draws: the next one is
    const turnPage = () => {
      if (shown()) pageEnded = true;
      else {
        blocks.length = 0;
        floats.length = 0;
      }
    };
    if (attrs.sourcePageBreakBefore === true || attrs.pageBreakBefore === true) {
      turnPage();
      if (pageEnded) return;
    }
    if (placed(node)) return;
    switch (type) {
      case "pageBreak":
        turnPage();
        return;
      case "sectionBreak":
        if (attrs.breakType !== "continuous") turnPage();
        return;
      case "heading": {
        const level = Number(attrs.level) || 1;
        const text = textOf(node);
        if (text) push({ tag: level <= 1 ? "h1" : level === 2 ? "h2" : "h3", text });
        for (const drawing of drawingsIn(node)) if (!placed(drawing)) inline(drawing);
        return;
      }
      case "paragraph": {
        const text = textOf(node);
        const drawings = drawingsIn(node);
        if (listKind) {
          if (text) push({ tag: listKind, text });
        } else if (text || !drawings.some((drawing) => !normalizeFloatAnchor(record(drawing.attrs)?.floatAnchor))) {
          // an empty paragraph keeps its line, as the page shows it
          push({ tag: "p", text });
        }
        for (const drawing of drawings) if (!placed(drawing)) inline(drawing);
        return;
      }
      case "bulletList":
        for (const child of children(node)) walk(child, "li");
        return;
      case "orderedList":
        for (const child of children(node)) walk(child, "oli");
        return;
      case "taskList":
        for (const child of children(node)) {
          const text = textOf(child);
          if (text) {
            push({
              tag: "task",
              text,
              checked: Boolean(record(child.attrs)?.checked),
            });
          }
        }
        return;
      case "blockquote": {
        const text = textOf(node);
        if (text) push({ tag: "quote", text });
        return;
      }
      case "codeBlock": {
        const text = textOf(node, 200);
        if (text) push({ tag: "code", text });
        return;
      }
      case "horizontalRule":
        push({ tag: "hr" });
        return;
      case "table": {
        const rows = children(node)
          .slice(0, MAX_TABLE_ROWS)
          .map((row) =>
            children(row)
              .slice(0, MAX_TABLE_COLS)
              .map((cell) => textOf(cell, 40)),
          );
        if (rows.length) push({ tag: "table", rows });
        return;
      }
      default:
        if (inline(node)) return;
        // pages, sections, columns, list items and other containers
        for (const child of children(node)) walk(child, listKind);
    }
  };
  if (root) walk(root, null);
  // trailing empty paragraphs say nothing
  while (blocks.length) {
    const last = blocks[blocks.length - 1]!;
    if (last.tag === "p" && !last.text) blocks.pop();
    else break;
  }
  floats.sort((a, b) => Number(b.behind) - Number(a.behind) || a.z - b.z);
  return { widthPx: geometry.pageWidthPx, heightPx: geometry.pageHeightPx, blocks, floats };
}

/** The first page's flowing blocks (see {@link documentThumbnailPage}). */
export function documentThumbnailBlocks(
  content: unknown,
  limit = 28,
  options: Omit<DocumentThumbnailOptions, "limit"> = {},
): DocumentThumbnailBlock[] {
  return documentThumbnailPage(content, { ...options, limit }).blocks;
}

export interface SpreadsheetThumbnailCell {
  text: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  background?: string;
  align?: "left" | "center" | "right";
}

export interface SpreadsheetThumbnailGrid {
  sheetName: string;
  rows: number;
  columns: number;
  /** Row-major, `rows` × `columns`; null for an empty cell. */
  cells: (SpreadsheetThumbnailCell | null)[][];
}

const HEX = /^#[0-9a-f]{3,8}$/i;

function cellText(cell: Record<string, unknown>): string {
  const value = cell.value ?? cell.display ?? cell.result;
  if (value === null || value === undefined) {
    return typeof cell.formula === "string" ? cell.formula : "";
  }
  if (typeof value === "number") {
    return Number.isInteger(value)
      ? String(value)
      : String(Number(value.toPrecision(10)));
  }
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value).slice(0, 60);
}

/** The top-left of the first sheet with any cells, with its cell styles. */
export function spreadsheetThumbnailGrid(
  content: unknown,
  maxRows = 18,
  maxColumns = 8,
): SpreadsheetThumbnailGrid | null {
  const root = record(content);
  const sheets = Array.isArray(root?.sheets) ? root.sheets : [];
  const sheet =
    sheets.map(record).find((candidate) => {
      const cells = record(candidate?.cells);
      return cells && Object.keys(cells).length > 0;
    }) ?? record(sheets[0]);
  if (!sheet) return null;
  const source = record(sheet.cells) ?? {};
  let usedRows = 0;
  let usedColumns = 0;
  const grid: (SpreadsheetThumbnailCell | null)[][] = Array.from(
    { length: maxRows },
    () => Array.from({ length: maxColumns }, () => null),
  );
  for (const [key, raw] of Object.entries(source)) {
    const match = /^(\d+):(\d+)$/.exec(key);
    const cell = record(raw);
    if (!match || !cell) continue;
    const row = Number(match[1]);
    const column = Number(match[2]);
    if (row >= maxRows || column >= maxColumns) continue;
    const style = record(cell.style) ?? {};
    const text = cellText(cell);
    const background =
      typeof style.backgroundColor === "string" && HEX.test(style.backgroundColor)
        ? style.backgroundColor
        : undefined;
    if (!text && !background) continue;
    const align = style.textAlign;
    grid[row]![column] = {
      text,
      bold: style.bold === true || undefined,
      italic: style.italic === true || undefined,
      color:
        typeof style.color === "string" && HEX.test(style.color)
          ? style.color
          : undefined,
      background,
      align:
        align === "center" || align === "right" || align === "left"
          ? align
          : typeof cell.value === "number"
            ? "right"
            : undefined,
    };
    usedRows = Math.max(usedRows, row + 1);
    usedColumns = Math.max(usedColumns, column + 1);
  }
  return {
    sheetName: typeof sheet.name === "string" ? sheet.name : "",
    // a little blank sheet around the data, as the grid shows it
    rows: Math.min(maxRows, Math.max(12, usedRows + 3)),
    columns: Math.min(maxColumns, Math.max(5, usedColumns + 1)),
    cells: grid,
  };
}

/** The first lines of a code page (a shared app's JSON is shown pretty). */
export function codeThumbnailLines(
  content: unknown,
  maxLines = 30,
): { language: string; lines: string[] } {
  const root = record(content);
  const language = typeof root?.language === "string" ? root.language : "";
  let text = typeof root?.content === "string" ? root.content : "";
  if (/^\s*[{[]/.test(text)) {
    try {
      text = JSON.stringify(JSON.parse(text), null, 2);
    } catch {
      // not JSON after all: shown as written
    }
  }
  const lines = text.split(/\r?\n/).slice(0, maxLines).map((line) => line.slice(0, 120));
  return { language, lines: lines.length === 1 && !lines[0] ? [] : lines };
}
