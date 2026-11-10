/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * 图表 on a PDF page (ONLYOFFICE's and Office's Insert Chart): a chart made in
 * MRAI, kept as one managed annotation (apps/api native_edit_primitives
 * `_upsert_managed_chart_annotation`): its model, the editors' chart V2 attrs
 * (the deck's and the document's own), and its picture, the chart engine's
 * inert SVG drawn at the frame's size, which the writer draws in vector. The
 * page keeps it as its directive (MRNOTE_NATIVE_CHART_V1) among the page's
 * annotations, as it keeps a table.
 */

export const MANAGED_NATIVE_CHART_PREFIX = "[[MRNOTE_NATIVE_CHART_V1|";
/**
 * 智能图形 (Insert › SmartArt) is a managed object of the same kind: its model
 * the Docs editor's SmartArt shape group (lib/office-smartart), which the
 * writer draws itself, by the painter the Docs PDF export draws it with
 * (`_upsert_managed_smartart_annotation`); its directive carries no picture.
 */
export const MANAGED_NATIVE_SMARTART_PREFIX = "[[MRNOTE_NATIVE_SMARTART_V1|";
/** The rows of the page's managed objects this module plans. */
export const MANAGED_NATIVE_CHART_PREFIXES: readonly string[] = Object.freeze([
  MANAGED_NATIVE_CHART_PREFIX, MANAGED_NATIVE_SMARTART_PREFIX,
]);
const MAX_MODEL_BYTES = 256 * 1024;
/** The engine's own bound on a picture (CHART_RENDER_MAX_SVG_BYTES). */
const MAX_SVG_BYTES = 2_000_000;
/** CSS pixels a PDF point takes at 100 %: the size a picture is drawn at. */
export const PDF_CHART_PX_PER_POINT = 96 / 72;

export interface PdfChart {
  /** A SmartArt graphic rather than a chart: `attrs` is its Docs shape group
   * and `svg` the picture the page paints of it (never written). */
  readonly kind?: "smartart";
  readonly id: string;
  /** One-based page number. */
  readonly page: number;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  /** The chart's model: chart V2 attrs (ChartBlockAttrsV2). */
  readonly attrs: Record<string, unknown>;
  /** The engine's picture of it at the frame's size; none until drawn. */
  readonly svg: string | null;
  readonly deleted?: false;
}

export type ManagedNativeChartRecord =
  | PdfChart
  | { readonly id: string; readonly page: number; readonly deleted: true; readonly kind?: "smartart" };

export interface ManagedNativeChartAnnotationShadow {
  id: string;
  text: string;
  page: number;
  rect: [number, number, number, number];
}

function base64Url(text: string): string {
  const bytes = new TextEncoder().encode(text);
  let binary = "";
  for (let index = 0; index < bytes.length; index += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(index, index + 0x8000));
  }
  return btoa(binary).replaceAll("+", "-").replaceAll("/", "_").replace(/=+$/, "");
}

function fromBase64Url(token: string): string | null {
  try {
    const binary = atob(token.replaceAll("-", "+").replaceAll("_", "/") + "=".repeat((4 - (token.length % 4)) % 4));
    return new TextDecoder("utf-8", { fatal: true }).decode(Uint8Array.from(binary, (ch) => ch.charCodeAt(0)));
  } catch {
    return null;
  }
}

function encodeChartId(id: string, kind: "chart" | "smartart" = "chart"): string | null {
  const trimmed = id.trim();
  if (!trimmed) return null;
  const encoded = encodeURIComponent(trimmed).replace(/[!'()*]/g,
    (character) => `%${character.charCodeAt(0).toString(16).toUpperCase().padStart(2, "0")}`);
  const token = `${kind}.${encoded}`;
  return token.length <= 384 ? token : null;
}

function decodeChartId(token: string, kind: "chart" | "smartart" = "chart"): string | null {
  if (!token.startsWith(`${kind}.`)) return null;
  try {
    return decodeURIComponent(token.slice(kind.length + 1)) || null;
  } catch {
    return null;
  }
}

function boxToken(value: number): string {
  return String(Number(Math.min(1, Math.max(0, value)).toFixed(6)));
}

/**
 * The accents a page's charts are painted against: those of the theme an AI
 * generated PDF was made in (its recipe's, `source.artifactStudio.recipe.theme`,
 * pdf_artifact_studio), which its charts were drawn with when it was made, as
 * a Docs chart's Change Colors resolves against its document's theme. Empty
 * for any other page (Office's default theme).
 */
export function pdfChartThemeAccents(content: unknown): readonly string[] {
  const source = content && typeof content === "object" ? (content as Record<string, unknown>).source : null;
  const studio = source && typeof source === "object" ? (source as Record<string, unknown>).artifactStudio : null;
  const recipe = studio && typeof studio === "object" ? (studio as Record<string, unknown>).recipe : null;
  const theme = recipe && typeof recipe === "object" ? (recipe as Record<string, unknown>).theme : null;
  const accents = theme && typeof theme === "object" ? (theme as Record<string, unknown>).accents : null;
  if (!Array.isArray(accents) || accents.length === 0) return [];
  const colors = accents.slice(0, 6);
  return colors.every((color) => typeof color === "string" && /^#[0-9A-Fa-f]{6}$/.test(color))
    ? Object.freeze(colors as string[]) : [];
}

/** The picture's size for a chart's frame on a page `pageWidth` x
 * `pageHeight` points: CSS pixels at 100 %. */
export function pdfChartPictureSize(
  chart: Pick<PdfChart, "width" | "height">,
  pageWidth: number,
  pageHeight: number,
): { width: number; height: number } {
  return {
    width: Math.max(40, Math.round(chart.width * pageWidth * PDF_CHART_PX_PER_POINT)),
    height: Math.max(30, Math.round(chart.height * pageHeight * PDF_CHART_PX_PER_POINT)),
  };
}

/** The chart's directive; null while it has no picture, or for one the
 * writer cannot take. */
export function managedNativeChartDirective(
  record: ManagedNativeChartRecord,
  rotation: 0 | 90 | 180 | 270 | "auto" = "auto",
): string | null {
  const smartArt = record.kind === "smartart";
  const prefix = smartArt ? MANAGED_NATIVE_SMARTART_PREFIX : MANAGED_NATIVE_CHART_PREFIX;
  const nativeId = encodeChartId(record.id, smartArt ? "smartart" : "chart");
  if (!nativeId || !Number.isInteger(record.page) || record.page < 1) return null;
  if (record.deleted) return `${prefix}delete|${nativeId}]]`;
  if (!smartArt && (!record.svg || new TextEncoder().encode(record.svg).length > MAX_SVG_BYTES)) return null;
  const model = JSON.stringify(record.attrs);
  if (new TextEncoder().encode(model).length > MAX_MODEL_BYTES) return null;
  const { x, y, width, height } = record;
  if (![x, y, width, height].every(Number.isFinite) || x < 0 || y < 0 || x >= 1 || y >= 1 ||
    width <= 0 || height <= 0 || x + width > 1 + 1e-6 || y + height > 1 + 1e-6) return null;
  const frame = `${nativeId}|${rotation}|${[x, y, width, height].map(boxToken).join(",")}|${base64Url(model)}`;
  return smartArt ? `${prefix}upsert|${frame}]]` : `${prefix}upsert|${frame}|${base64Url(record.svg!)}]]`;
}

const DIRECTIVE_RE = /^\[\[MRNOTE_NATIVE_CHART_V1\|(upsert|delete)\|(chart\.[A-Za-z0-9._~%-]{1,378})(?:\|(auto|0|90|180|270)\|([0-9.]{1,20}),([0-9.]{1,20}),([0-9.]{1,20}),([0-9.]{1,20})\|([A-Za-z0-9_-]{4,400000})\|([A-Za-z0-9_-]{4,2800000}))?\]\]$/;

const SMARTART_DIRECTIVE_RE = /^\[\[MRNOTE_NATIVE_SMARTART_V1\|(upsert|delete)\|(smartart\.[A-Za-z0-9._~%-]{1,375})(?:\|(auto|0|90|180|270)\|([0-9.]{1,20}),([0-9.]{1,20}),([0-9.]{1,20}),([0-9.]{1,20})\|([A-Za-z0-9_-]{4,400000}))?\]\]$/;

function parseManagedNativeSmartArtShadow(text: string, page: number): ManagedNativeChartRecord | null {
  const match = SMARTART_DIRECTIVE_RE.exec(text);
  if (!match) return null;
  const id = decodeChartId(match[2], "smartart");
  if (!id) return null;
  if (match[1] === "delete") return match[3] === undefined ? { kind: "smartart", id, page, deleted: true } : null;
  if (match[3] === undefined) return null;
  const model = fromBase64Url(match[8]);
  if (model === null) return null;
  let attrs: unknown;
  try {
    attrs = JSON.parse(model);
  } catch {
    return null;
  }
  if (!attrs || typeof attrs !== "object" || Array.isArray(attrs)) return null;
  const [x, y, width, height] = [match[4], match[5], match[6], match[7]].map(Number);
  return { kind: "smartart", id, page, x, y, width, height, attrs: attrs as Record<string, unknown>, svg: null };
}

/** A page annotation row back as the chart (or SmartArt) it is (null for any other). */
export function parseManagedNativeChartShadow(value: unknown): ManagedNativeChartRecord | null {
  if (!value || typeof value !== "object") return null;
  const row = value as Record<string, unknown>;
  const page = Number(row.page);
  if (!Number.isInteger(page) || page < 1 || page > 10_000) return null;
  const text = typeof row.text === "string" ? row.text : "";
  if (text.startsWith(MANAGED_NATIVE_SMARTART_PREFIX)) return parseManagedNativeSmartArtShadow(text, page);
  const match = DIRECTIVE_RE.exec(text);
  if (!match) return null;
  const id = decodeChartId(match[2]);
  if (!id) return null;
  if (match[1] === "delete") return match[3] === undefined ? { id, page, deleted: true } : null;
  if (match[3] === undefined) return null;
  const model = fromBase64Url(match[8]);
  const svg = fromBase64Url(match[9]);
  if (model === null || svg === null) return null;
  let attrs: unknown;
  try {
    attrs = JSON.parse(model);
  } catch {
    return null;
  }
  if (!attrs || typeof attrs !== "object" || Array.isArray(attrs)) return null;
  const [x, y, width, height] = [match[4], match[5], match[6], match[7]].map(Number);
  return { id, page, x, y, width, height, attrs: attrs as Record<string, unknown>, svg };
}

export interface ManagedNativeChartPlan {
  annotations: ManagedNativeChartAnnotationShadow[];
  supported: boolean;
  reason?: "invalid_chart" | "duplicate_id";
}

/** Every chart with a picture as page annotation rows (one still being
 * drawn waits for its picture); one the writer cannot take fails the plan. */
export function planManagedNativeChartAnnotations(
  records: readonly ManagedNativeChartRecord[],
  rotations: ReadonlyMap<number, 0 | 90 | 180 | 270> = new Map(),
): ManagedNativeChartPlan {
  const ids = new Set<string>();
  const annotations: ManagedNativeChartAnnotationShadow[] = [];
  for (const record of records) {
    if (ids.has(record.id)) return { annotations: [], supported: false, reason: "duplicate_id" };
    ids.add(record.id);
    if (!record.deleted && !record.svg && record.kind !== "smartart") continue;
    const text = managedNativeChartDirective(record, rotations.get(record.page) ?? "auto");
    if (!text) return { annotations: [], supported: false, reason: "invalid_chart" };
    annotations.push({ id: `__mrnote_native_chart__${record.id}`, text, page: record.page, rect: [24, 24, 44, 44] });
  }
  return { annotations, supported: true };
}
