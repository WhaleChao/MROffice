/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Draw ▸ Ink to Math, the editors' side: the strokes of the Math Input
 * Control rasterised to one PNG and sent to the workspace's vision model,
 * which answers LaTeX (`apps/api/app/routers/ink_math.py`). Shared by the
 * Docs, deck and worksheet Draw tabs.
 */
import { apiGet, apiPostFormData } from "./api";

/** The viewer's one-time consent to send handwriting to the workspace's AI service. */
export const INK_MATH_PRIVACY_STORAGE_KEY = "mrnote.ink-math.privacy-acknowledged";

export function inkMathPrivacyAcknowledged(): boolean {
  try {
    return typeof window !== "undefined" && window.localStorage.getItem(INK_MATH_PRIVACY_STORAGE_KEY) === "true";
  } catch {
    return false;
  }
}

export interface InkMathPoint {
  readonly x: number;
  readonly y: number;
}

export interface InkMathCapability {
  readonly available: boolean;
  readonly reasonCode: string | null;
  /**
   * Who reads the strokes: the server's built-in readers ("local", nothing
   * leaves the server) or the workspace's AI vision model ("vision").
   */
  readonly engine?: "local" | "vision";
  /** Whether the handwriting is sent to a configured AI service, so the one-time consent applies. */
  readonly needsAcknowledgement?: boolean;
  readonly modelId: string | null;
  readonly residencyRegion: string | null;
  readonly maxImageBytes: number;
  readonly acceptedMediaTypes: readonly string[];
}

export interface InkMathRecognition {
  readonly latex: string;
  readonly alternatives: readonly string[];
  readonly confidence: number;
  readonly modelId: string;
  readonly policyVersion: number;
  readonly policyMarkerSha256: string;
}

/**
 * Whether this capability answer means the handwriting is sent to an AI
 * service the reader must agree to once; the server's built-in readers
 * ("local") send nothing on and ask nothing.
 */
export function officeInkReadingNeedsConsent(
  capability: Pick<InkMathCapability, "needsAcknowledgement" | "engine">,
): boolean {
  if (typeof capability.needsAcknowledgement === "boolean") return capability.needsAcknowledgement;
  return capability.engine !== "local";
}

export async function loadInkMathCapability(pageId: string, signal?: AbortSignal): Promise<InkMathCapability> {
  return apiGet<InkMathCapability>(`/api/v1/pages/${encodeURIComponent(pageId)}/ink-math/capability`, { signal });
}

/** Side of the PNG the strokes are drawn into (the model reads it at this size). */
export const INK_MATH_RASTER_MAX_SIDE = 1600;
const INK_MATH_RASTER_PADDING = 24;

/**
 * The strokes as a PNG: black ink on white, padded, the long side at most
 * {@link INK_MATH_RASTER_MAX_SIDE} px, line width scaled with the drawing so a
 * small equation is not read as hairlines. Null when there is nothing drawn
 * or the browser has no canvas.
 */
export async function rasterizeInkMathStrokes(
  strokes: readonly (readonly InkMathPoint[])[],
  options: { readonly lineWidth?: number } = {},
): Promise<Blob | null> {
  const points = strokes.flat();
  if (points.length === 0 || typeof document === "undefined") return null;
  let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
  for (const point of points) {
    minX = Math.min(minX, point.x); minY = Math.min(minY, point.y);
    maxX = Math.max(maxX, point.x); maxY = Math.max(maxY, point.y);
  }
  const width = Math.max(1, maxX - minX);
  const height = Math.max(1, maxY - minY);
  const scale = Math.min(1, (INK_MATH_RASTER_MAX_SIDE - 2 * INK_MATH_RASTER_PADDING) / Math.max(width, height));
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(32, Math.ceil(width * scale) + 2 * INK_MATH_RASTER_PADDING);
  canvas.height = Math.max(32, Math.ceil(height * scale) + 2 * INK_MATH_RASTER_PADDING);
  const context = canvas.getContext("2d");
  if (!context) return null;
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, canvas.width, canvas.height);
  context.strokeStyle = "#000000";
  context.lineCap = "round";
  context.lineJoin = "round";
  context.lineWidth = Math.max(2, (options.lineWidth ?? 3) * Math.max(scale, 0.5));
  for (const stroke of strokes) {
    if (stroke.length === 0) continue;
    context.beginPath();
    stroke.forEach((point, index) => {
      const x = (point.x - minX) * scale + INK_MATH_RASTER_PADDING;
      const y = (point.y - minY) * scale + INK_MATH_RASTER_PADDING;
      if (index === 0) context.moveTo(x, y);
      else context.lineTo(x, y);
    });
    if (stroke.length === 1) {
      const x = (stroke[0]!.x - minX) * scale + INK_MATH_RASTER_PADDING;
      const y = (stroke[0]!.y - minY) * scale + INK_MATH_RASTER_PADDING;
      context.lineTo(x + 0.1, y + 0.1);
    }
    context.stroke();
  }
  return new Promise((resolve) => canvas.toBlob((blob) => resolve(blob), "image/png"));
}

export async function recognizeInkMath(
  pageId: string,
  image: Blob,
  options: { readonly privacyAcknowledged: boolean; readonly signal?: AbortSignal },
): Promise<InkMathRecognition> {
  const form = new FormData();
  form.append("file", image, "ink-math.png");
  form.append("privacy_acknowledged", options.privacyAcknowledged ? "true" : "false");
  return apiPostFormData<InkMathRecognition>(
    `/api/v1/pages/${encodeURIComponent(pageId)}/ink-math/recognize`,
    form,
    { signal: options.signal },
  );
}
