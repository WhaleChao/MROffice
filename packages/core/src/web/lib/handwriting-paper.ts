/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Handwriting documents (手写文档): a Docs page set up for the pen, with a
 * cover (the notebook cover system: lib/notebook-cover + Book3D) and a paper
 * every page is drawn on.
 *
 * Each paper is ONE parametric definition (handwriting-paper-catalog.json):
 * a list of primitives in millimetres on the page, measured from an edge
 * (a negative number counts from the far edge, "N%" is a share of the side).
 * The same definition draws the gallery's thumbnail, the editor's pages, the
 * printed PDF (the pages print themselves) and, through the API's twin
 * (app/services/handwriting_paper.py reading a byte-identical copy of the
 * catalog), the picture a .docx carries behind the text in its header, the
 * way Word keeps a picture watermark.
 *
 * The setting lives on the document root as `handwritingDocument`; the page
 * size and orientation are the document's own `pageSize` / `pageOrientation`,
 * so the pages, the print and the export already follow them.
 */

import catalog from "./handwriting-paper-catalog.json";
import {
  PAGE_SIZE_PRESET_VALUES,
  PAGE_SIZE_TWIPS,
  type PageSize,
} from "../components/console/editor/native-document/pageSetup";

export type HandwritingLength = number | string;

export type HandwritingTone = "rule" | "strong" | "accent" | "band" | "dot";

export type HandwritingPrimitive =
  | { t: "rules"; x0: HandwritingLength; x1: HandwritingLength; y0: HandwritingLength; y1: HandwritingLength; pitch: number; tone: HandwritingTone; w: number; dash?: number[] }
  | { t: "cols"; x0: HandwritingLength; x1: HandwritingLength; y0: HandwritingLength; y1: HandwritingLength; pitch: number; tone: HandwritingTone; w: number; dash?: number[] }
  // `snap`: the grid stops at a whole number of this many millimetres (a
  // graph paper's fine grid ends on its coarse grid's last line).
  | { t: "grid"; x0: HandwritingLength; x1: HandwritingLength; y0: HandwritingLength; y1: HandwritingLength; pitch: number; snap?: number; tone: HandwritingTone; w: number }
  | { t: "dots"; x0: HandwritingLength; x1: HandwritingLength; y0: HandwritingLength; y1: HandwritingLength; pitch: number; pitchY?: number; stagger?: boolean; r: number; tone: HandwritingTone }
  | { t: "line"; x0: HandwritingLength; x1: HandwritingLength; y0: HandwritingLength; y1: HandwritingLength; tone: HandwritingTone; w: number; dash?: number[] }
  | { t: "box"; x0: HandwritingLength; x1: HandwritingLength; y0: HandwritingLength; y1: HandwritingLength; fill?: HandwritingTone; stroke?: HandwritingTone; w?: number; r?: number }
  | { t: "cells"; x0: HandwritingLength; x1: HandwritingLength; y0: HandwritingLength; y1: HandwritingLength; cols: number; rows: number; tone: HandwritingTone; w: number; head?: number }
  | { t: "staves"; x0: HandwritingLength; x1: HandwritingLength; y0: HandwritingLength; y1: HandwritingLength; lines: number; gap: number; space: number; tone: HandwritingTone; w: number }
  | { t: "checks"; x: HandwritingLength; y0: HandwritingLength; y1: HandwritingLength; pitch: number; size: number; tone: HandwritingTone; w: number };

interface Catalog {
  version: number;
  colors: Record<string, Record<"paper" | HandwritingTone, string>>;
  categories: { id: string; papers: string[] }[];
  papers: Record<string, HandwritingPrimitive[]>;
}

const CATALOG = catalog as unknown as Catalog;

export const HANDWRITING_PAPER_CATEGORIES = CATALOG.categories.map((category) => ({
  id: category.id as HandwritingPaperCategoryId,
  papers: category.papers as HandwritingPaperId[],
}));
export type HandwritingPaperCategoryId = "basic" | "writing" | "planning" | "music";
export type HandwritingPaperId =
  | "blank" | "dotGrid" | "narrowRuled" | "wideRuled" | "squared" | "graph" | "isometricDots"
  | "cornell" | "legalPad" | "singleColumn" | "singleColumnMixed" | "twoColumn" | "threeColumn"
  | "accounting" | "monthlyPlanner" | "weeklyPlanner" | "dailyPlanner" | "todoList"
  | "staff" | "guitarTab";
export const HANDWRITING_PAPER_IDS = Object.keys(CATALOG.papers) as HandwritingPaperId[];

export type HandwritingPaperColorId = "white" | "cream" | "yellow" | "mist" | "dark";
export const HANDWRITING_PAPER_COLOR_IDS = Object.keys(CATALOG.colors) as HandwritingPaperColorId[];

/** The page sizes a handwriting document offers (Docs' own presets). */
/** The Docs page sizes (its Layout > Size gallery), one list for both. */
export const HANDWRITING_PAGE_SIZES = PAGE_SIZE_PRESET_VALUES;
export type HandwritingPageSize = Exclude<PageSize, "custom">;
export type HandwritingOrientation = "portrait" | "landscape";

/** Paper sizes in millimetres (portrait), from the twips Word writes. */
const PAGE_SIZE_MM = Object.fromEntries(
  PAGE_SIZE_PRESET_VALUES.map((size) => [
    size,
    { width: (PAGE_SIZE_TWIPS[size].width / 1440) * 25.4, height: (PAGE_SIZE_TWIPS[size].height / 1440) * 25.4 },
  ]),
) as Record<HandwritingPageSize, { width: number; height: number }>;

export function handwritingPageMm(
  size: string | null | undefined,
  orientation: string | null | undefined,
): { width: number; height: number } {
  const base = PAGE_SIZE_MM[(HANDWRITING_PAGE_SIZES as readonly string[]).includes(String(size)) ? (size as HandwritingPageSize) : "a4"];
  return orientation === "landscape" ? { width: base.height, height: base.width } : { ...base };
}

/** The notebook cover a handwriting document wears (lib/notebook-cover's CoverPreference). */
export interface HandwritingCover {
  template: string;
  color: string;
  spineColor: string;
  textTheme: "light" | "dark";
  imageUrl?: string;
  colorPicked?: boolean | null;
}

export interface HandwritingDocumentSetting {
  version: 1;
  paper: HandwritingPaperId;
  paperColor: HandwritingPaperColorId;
  cover: HandwritingCover | null;
}

const HEX = /^#[0-9a-fA-F]{6}$/;
const TEMPLATE = /^[A-Za-z][A-Za-z0-9_-]{0,63}$/;
const DATA_IMAGE = /^data:image\/(?:png|jpe?g|webp);base64,[A-Za-z0-9+/=\s]+$/i;
/** A cover picture stated as a data URL: the notebook cover's 2 MiB, base64. */
const MAX_DATA_IMAGE_CHARS = 2_800_000;

function coverImageUrl(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  if (!trimmed) return undefined;
  if (trimmed.toLowerCase().startsWith("data:")) {
    return trimmed.length <= MAX_DATA_IMAGE_CHARS && DATA_IMAGE.test(trimmed) ? trimmed : undefined;
  }
  if (trimmed.length > 2048 || !/^https:\/\/[^\s"'<>()\\]+$/i.test(trimmed)) return undefined;
  return trimmed;
}

export function normalizeHandwritingCover(value: unknown): HandwritingCover | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const raw = value as Record<string, unknown>;
  if (typeof raw.template !== "string" || !TEMPLATE.test(raw.template)) return null;
  if (typeof raw.color !== "string" || !HEX.test(raw.color)) return null;
  if (typeof raw.spineColor !== "string" || !HEX.test(raw.spineColor)) return null;
  if (raw.textTheme !== "light" && raw.textTheme !== "dark") return null;
  const cover: HandwritingCover = {
    template: raw.template,
    color: raw.color.toLowerCase(),
    spineColor: raw.spineColor.toLowerCase(),
    textTheme: raw.textTheme,
  };
  const imageUrl = coverImageUrl(raw.imageUrl);
  if (imageUrl) cover.imageUrl = imageUrl;
  if (raw.colorPicked === true) cover.colorPicked = true;
  return cover;
}

/** The stored setting, or null when the page is not a handwriting document. */
export function normalizeHandwritingDocument(value: unknown): HandwritingDocumentSetting | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const raw = value as Record<string, unknown>;
  if (raw.version !== 1) return null;
  if (typeof raw.paper !== "string" || !Object.hasOwn(CATALOG.papers, raw.paper)) return null;
  if (typeof raw.paperColor !== "string" || !Object.hasOwn(CATALOG.colors, raw.paperColor)) return null;
  return {
    version: 1,
    paper: raw.paper as HandwritingPaperId,
    paperColor: raw.paperColor as HandwritingPaperColorId,
    cover: normalizeHandwritingCover(raw.cover),
  };
}

export function handwritingPaperColors(color: HandwritingPaperColorId) {
  return CATALOG.colors[color] ?? CATALOG.colors.white!;
}

export function handwritingPaperIsDark(color: HandwritingPaperColorId): boolean {
  return color === "dark";
}

/** Each channel at or under this is a black ink (#000000 to #404040). */
const NEAR_BLACK_CHANNEL = 0x40;








export function handwritingInkColor(color: string, paperColor: HandwritingPaperColorId | null | undefined): string {
  return paperColor && handwritingPaperIsDark(paperColor) ? handwritingInkColorOnDark(color) : color;
}






export function handwritingCoverIsDark(cover: HandwritingCover | null | undefined): boolean {
  return cover?.textTheme === "light";
}

/** Whether the page at `index` (0 is the cover when there is one) is dark under the pen. */
export function handwritingPageIsDark(setting: HandwritingDocumentSetting | null | undefined, index: number): boolean {
  if (!setting) return false;
  if (setting.cover && index === 0) return handwritingCoverIsDark(setting.cover);
  return handwritingPaperIsDark(setting.paperColor);
}

/** The colour a pen writes in on a dark page: a near-black ink is white; any other is its own. */
export function handwritingInkColorOnDark(color: string): string {
  const hex = /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(color);
  if (!hex) return color;
  const nearBlack = hex.slice(1).every((part) => Number.parseInt(part, 16) <= NEAR_BLACK_CHANNEL);
  return nearBlack ? "#ffffff" : color;
}

/* ------------------------------------------------------------ drawing */

function at(value: HandwritingLength, side: number): number {
  if (typeof value === "string") {
    const share = Number.parseFloat(value);
    return Number.isFinite(share) ? (share / 100) * side : 0;
  }
  return value < 0 ? side + value : value;
}

function n(value: number): string {
  const rounded = Math.round(value * 100) / 100;
  return Object.is(rounded, -0) ? "0" : String(rounded);
}

const EPSILON = 1e-6;
const PX_PER_MM = 96 / 25.4;

function stroke(color: string, width: number, dash?: number[]): string {
  return `fill="none" stroke="${color}" stroke-width="${n(width)}"${
    dash && dash.length ? ` stroke-dasharray="${dash.map(n).join(" ")}"` : ""
  }`;
}

/**
 * The paper as an SVG document: `widthMm` x `heightMm` in user units, sized
 * in CSS pixels at 96 dpi (so its intrinsic ratio is the page's).
 * `minStrokeMm` keeps a thumbnail's hairlines visible at its size (a page is
 * drawn at 0: the definition's own weights).
 */
export function handwritingPaperSvg(
  paper: string,
  color: HandwritingPaperColorId,
  widthMm: number,
  heightMm: number,
  minStrokeMm = 0,
): string {
  const line = (tone: string, width: number, dash?: number[]) =>
    stroke(tone, Math.max(width, minStrokeMm), dash);
  const tones = handwritingPaperColors(color);
  const W = widthMm;
  const H = heightMm;
  const parts: string[] = [`<rect width="${n(W)}" height="${n(H)}" fill="${tones.paper}"/>`];
  const defs: string[] = [];
  const primitives = CATALOG.papers[paper] ?? [];
  primitives.forEach((primitive, index) => {
    switch (primitive.t) {
      case "rules":
      case "cols": {
        const x0 = at(primitive.x0, W);
        const x1 = at(primitive.x1, W);
        const y0 = at(primitive.y0, H);
        const y1 = at(primitive.y1, H);
        const d: string[] = [];
        if (primitive.t === "rules") {
          for (let y = y0; y <= y1 + EPSILON; y += primitive.pitch) d.push(`M${n(x0)} ${n(y)}H${n(x1)}`);
        } else {
          for (let x = x0; x <= x1 + EPSILON; x += primitive.pitch) d.push(`M${n(x)} ${n(y0)}V${n(y1)}`);
        }
        if (d.length) parts.push(`<path d="${d.join("")}" ${line(tones[primitive.tone], primitive.w, primitive.dash)}/>`);
        break;
      }
      case "grid": {
        const x0 = at(primitive.x0, W);
        const y0 = at(primitive.y0, H);
        let across = Math.floor((at(primitive.x1, W) - x0) / primitive.pitch + EPSILON);
        let down = Math.floor((at(primitive.y1, H) - y0) / primitive.pitch + EPSILON);
        if (primitive.snap && primitive.snap > primitive.pitch) {
          // A graph paper's fine grid ran on past its coarse grid's last line
          // (6 mm on A4) where the page did not hold whole coarse cells.
          const cells = Math.round(primitive.snap / primitive.pitch);
          across = Math.floor(across / cells) * cells;
          down = Math.floor(down / cells) * cells;
        }
        if (across < 1 || down < 1) break;
        const xe = x0 + across * primitive.pitch;
        const ye = y0 + down * primitive.pitch;
        const d: string[] = [];
        for (let i = 0; i <= across; i += 1) d.push(`M${n(x0 + i * primitive.pitch)} ${n(y0)}V${n(ye)}`);
        for (let j = 0; j <= down; j += 1) d.push(`M${n(x0)} ${n(y0 + j * primitive.pitch)}H${n(xe)}`);
        parts.push(`<path d="${d.join("")}" ${line(tones[primitive.tone], primitive.w)}/>`);
        break;
      }
      case "dots": {
        const x0 = at(primitive.x0, W);
        const y0 = at(primitive.y0, H);
        const pitchY = primitive.pitchY ?? primitive.pitch;
        const across = Math.floor((at(primitive.x1, W) - x0) / primitive.pitch + EPSILON) + 1;
        const down = Math.floor((at(primitive.y1, H) - y0) / pitchY + EPSILON) + 1;
        if (across < 1 || down < 1) break;
        const id = `p${index}`;
        const fill = tones[primitive.tone];
        if (primitive.stagger) {
          const rows = Math.ceil(down / 2);
          defs.push(
            `<pattern id="${id}" patternUnits="userSpaceOnUse" x="${n(x0 - primitive.pitch / 4)}" y="${n(y0 - pitchY / 2)}" width="${n(primitive.pitch)}" height="${n(pitchY * 2)}">` +
              `<circle cx="${n(primitive.pitch / 4)}" cy="${n(pitchY / 2)}" r="${n(Math.max(primitive.r, minStrokeMm * 0.75))}" fill="${fill}"/>` +
              `<circle cx="${n((primitive.pitch * 3) / 4)}" cy="${n(pitchY * 1.5)}" r="${n(Math.max(primitive.r, minStrokeMm * 0.75))}" fill="${fill}"/></pattern>`,
          );
          parts.push(
            `<rect x="${n(x0 - primitive.pitch / 4)}" y="${n(y0 - pitchY / 2)}" width="${n(across * primitive.pitch)}" height="${n(Math.min(down, rows * 2) * pitchY)}" fill="url(#${id})"/>`,
          );
        } else {
          defs.push(
            `<pattern id="${id}" patternUnits="userSpaceOnUse" x="${n(x0 - primitive.pitch / 2)}" y="${n(y0 - pitchY / 2)}" width="${n(primitive.pitch)}" height="${n(pitchY)}">` +
              `<circle cx="${n(primitive.pitch / 2)}" cy="${n(pitchY / 2)}" r="${n(Math.max(primitive.r, minStrokeMm * 0.75))}" fill="${fill}"/></pattern>`,
          );
          parts.push(
            `<rect x="${n(x0 - primitive.pitch / 2)}" y="${n(y0 - pitchY / 2)}" width="${n(across * primitive.pitch)}" height="${n(down * pitchY)}" fill="url(#${id})"/>`,
          );
        }
        break;
      }
      case "line": {
        parts.push(
          `<path d="M${n(at(primitive.x0, W))} ${n(at(primitive.y0, H))}L${n(at(primitive.x1, W))} ${n(at(primitive.y1, H))}" ${line(tones[primitive.tone], primitive.w, primitive.dash)}/>`,
        );
        break;
      }
      case "box": {
        const x0 = at(primitive.x0, W);
        const y0 = at(primitive.y0, H);
        const width = at(primitive.x1, W) - x0;
        const height = at(primitive.y1, H) - y0;
        if (width <= 0 || height <= 0) break;
        const fill = primitive.fill ? tones[primitive.fill] : "none";
        const outline = primitive.stroke
          ? ` stroke="${tones[primitive.stroke]}" stroke-width="${n(Math.max(primitive.w ?? 0.25, minStrokeMm))}"`
          : "";
        parts.push(
          `<rect x="${n(x0)}" y="${n(y0)}" width="${n(width)}" height="${n(height)}"${primitive.r ? ` rx="${n(primitive.r)}"` : ""} fill="${fill}"${outline}/>`,
        );
        break;
      }
      case "cells": {
        const x0 = at(primitive.x0, W);
        const y0 = at(primitive.y0, H);
        const x1 = at(primitive.x1, W);
        const y1 = at(primitive.y1, H);
        if (x1 <= x0 || y1 <= y0) break;
        const cw = (x1 - x0) / primitive.cols;
        const ch = (y1 - y0) / primitive.rows;
        const d: string[] = [`M${n(x0)} ${n(y0)}H${n(x1)}V${n(y1)}H${n(x0)}Z`];
        for (let i = 1; i < primitive.cols; i += 1) d.push(`M${n(x0 + i * cw)} ${n(y0)}V${n(y1)}`);
        for (let j = 1; j < primitive.rows; j += 1) d.push(`M${n(x0)} ${n(y0 + j * ch)}H${n(x1)}`);
        parts.push(`<path d="${d.join("")}" ${line(tones[primitive.tone], primitive.w)}/>`);
        if (primitive.head && primitive.head < ch) {
          const heads: string[] = [];
          for (let j = 0; j < primitive.rows; j += 1) heads.push(`M${n(x0)} ${n(y0 + j * ch + primitive.head)}H${n(x1)}`);
          parts.push(`<path d="${heads.join("")}" ${line(tones.rule, primitive.w * 0.8)}/>`);
        }
        break;
      }
      case "staves": {
        const x0 = at(primitive.x0, W);
        const x1 = at(primitive.x1, W);
        const y1 = at(primitive.y1, H);
        const height = (primitive.lines - 1) * primitive.gap;
        const d: string[] = [];
        for (let top = at(primitive.y0, H); top + height <= y1 + EPSILON; top += height + primitive.space) {
          for (let i = 0; i < primitive.lines; i += 1) d.push(`M${n(x0)} ${n(top + i * primitive.gap)}H${n(x1)}`);
          d.push(`M${n(x0)} ${n(top)}V${n(top + height)}M${n(x1)} ${n(top)}V${n(top + height)}`);
        }
        if (d.length) parts.push(`<path d="${d.join("")}" ${line(tones[primitive.tone], primitive.w)}/>`);
        break;
      }
      case "checks": {
        const x = at(primitive.x, W);
        const y1 = at(primitive.y1, H);
        const d: string[] = [];
        for (let y = at(primitive.y0, H); y <= y1 + EPSILON; y += primitive.pitch) {
          const top = y - primitive.size - 1.2;
          if (top < 0) continue;
          d.push(`M${n(x)} ${n(top)}h${n(primitive.size)}v${n(primitive.size)}h${n(-primitive.size)}Z`);
        }
        if (d.length) parts.push(`<path d="${d.join("")}" ${line(tones[primitive.tone], primitive.w)}/>`);
        break;
      }
      default:
        break;
    }
  });
  return (
    `<svg xmlns="http://www.w3.org/2000/svg" width="${n(W * PX_PER_MM)}" height="${n(H * PX_PER_MM)}" viewBox="0 0 ${n(W)} ${n(H)}" preserveAspectRatio="none">` +
    (defs.length ? `<defs>${defs.join("")}</defs>` : "") +
    parts.join("") +
    "</svg>"
  );
}

const dataUrlCache = new Map<string, string>();

/** The paper as a CSS-ready data URL (remembered per paper, colour and size). */
export function handwritingPaperDataUrl(
  paper: string,
  color: HandwritingPaperColorId,
  widthMm: number,
  heightMm: number,
  minStrokeMm = 0,
): string {
  const key = `${paper}|${color}|${n(widthMm)}|${n(heightMm)}|${n(minStrokeMm)}`;
  const kept = dataUrlCache.get(key);
  if (kept) return kept;
  const url = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(handwritingPaperSvg(paper, color, widthMm, heightMm, minStrokeMm))}`;
  if (dataUrlCache.size > 160) dataUrlCache.clear();
  dataUrlCache.set(key, url);
  return url;
}

/* ------------------------------------------------------------ creating */

export interface HandwritingDocumentDraft {
  paper: HandwritingPaperId;
  paperColor: HandwritingPaperColorId;
  pageSize: HandwritingPageSize;
  orientation: HandwritingOrientation;
  cover: HandwritingCover | null;
}

/**
 * A new handwriting document's content: the setting, Docs' page setup for
 * the chosen paper, and its pages -- the cover first (a page of its own,
 * ended by a page break) when it has one, then a page of paper.
 */
export function handwritingDocumentContent(draft: HandwritingDocumentDraft): {
  type: "doc";
  attrs: Record<string, unknown>;
  content: Array<{ type: string }>;
} {
  const setting: HandwritingDocumentSetting = {
    version: 1,
    paper: draft.paper,
    paperColor: draft.paperColor,
    cover: draft.cover ? normalizeHandwritingCover(draft.cover) : null,
  };
  return {
    type: "doc",
    attrs: {
      handwritingDocument: setting,
      pageSize: draft.pageSize,
      pageOrientation: draft.orientation,
      pageMargins: "narrow",
    },
    content: setting.cover
      ? [{ type: "paragraph" }, { type: "pageBreak" }, { type: "paragraph" }]
      : [{ type: "paragraph" }],
  };
}
