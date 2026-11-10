/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Video Format > Video Border: a video is a `p:pic` and its border is the
 * picture's own line (`p:spPr/a:ln`; Microsoft Support, "Add, format, and
 * record video"). The embed block keeps it in the picture border fields
 * (`borderColor` / `borderWidth` / `borderStyle`, px), read, written and
 * patched by the API exactly as a picture's are. Pure: the window only wires.
 */

import type { CSSProperties } from "react";

import type { PresentationEmbedBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import {
  normalizePictureClipShape,
  pictureCropShapePreset,
} from "./presentation-picture-crop-shape";
import { presentationShapeCatalogEntryForPreset } from "@mrnote/core/web/lib/presentation-shape-catalog.ts";
import type { ShapeKind } from "@mrnote/core/web/lib/presentation-shapes.ts";
import {
  isOfficeRecolorPreset,
  officeAdjustFilterSteps,
  type OfficeAdjustFilterStep,
  type OfficeRecolorPresetId,
} from "@mrnote/core/web/lib/office-recolor-presets.ts";
import type { OfficeThemeGridColors } from "@mrnote/core/web/lib/office-theme-color-grid.ts";
import { officeVideoStyle, type OfficeVideoStyleId } from "@mrnote/core/web/lib/office-video-styles/catalog.ts";
import { isVideoPosterSource, MAX_VIDEO_POSTER_SRC_LENGTH } from "@mrnote/core/web/lib/presentation-video-poster-source.ts";
import {
  DEFAULT_IMAGE_BORDER_WIDTH,
  normalizeImageBorderColor,
  normalizeImageBorderWidth,
  type ImageBorderStyle,
} from "@mrnote/core/web/lib/presentation-image-effects.ts";

export interface VideoBorderPayload {
  readonly style: ImageBorderStyle;
  readonly color?: string;
  readonly width?: number;
}

const BORDER_KEYS = ["borderColor", "borderWidth", "borderStyle"] as const;

/** Whether the block is a video that can carry a border (the file's own line is kept). */
export function videoBorderEditable(block: PresentationEmbedBlock | null | undefined): boolean {
  if (!block || block.kind !== "video") return false;
  const source = block.source as Record<string, unknown> | undefined;
  
  // PowerPoint keeps Video Border available after a style.
  if ((block as Record<string, unknown>).videoStyle !== undefined)
    return source?.videoStyleDisposition !== "preserve-only";
  return source?.pictureBorderDisposition !== "preserve-only";
}







export function videoStyleOutline(block: {
  readonly videoStyle?: unknown;
  readonly borderColor?: unknown;
  readonly borderWidth?: unknown;
  readonly borderStyle?: unknown;
}): { color: string; widthPx: number; dash: "solid" | "dashed" | "dotted" } | "none" | null {
  const style = officeVideoStyle(block.videoStyle);
  if (!style) return null;
  const line = style.line && Number(style.line.width) > 0 ? style.line : null;
  if (block.borderStyle === "none") return line ? "none" : null;
  const widthPx = normalizeImageBorderWidth(block.borderWidth);
  const color = normalizeImageBorderColor(block.borderColor, null);
  if (!widthPx && !color && block.borderStyle === undefined) return null;
  const outline: { color: string; widthPx: number; dash: "solid" | "dashed" | "dotted" } = {
    color: color ?? line?.color ?? "#000000",
    widthPx: widthPx || (line ? Number(line.width) / 9525 : 1),
    dash:
      block.borderStyle === "dashed" || block.borderStyle === "dotted"
        ? block.borderStyle
        : ("solid" as const),
  };
  return line &&
    !line.gradient &&
    outline.dash === "solid" &&
    outline.color.toLowerCase() === line.color?.toLowerCase() &&
    Math.round(outline.widthPx * 9525) === Math.round(Number(line.width))
    ? null
    : outline;
}

/** The video's current border, as the picture border menu shows one. */
export function videoBorderState(block: PresentationEmbedBlock): {
  visible: boolean;
  color: string;
  width: number;
  style: Exclude<ImageBorderStyle, "none">;
} {
  const record = block as Record<string, unknown>;
  const videoStyle = officeVideoStyle(record.videoStyle);
  if (videoStyle) {
    // The style's own line until a border is set over it.
    const outline = videoStyleOutline(record);
    const line = videoStyle.line && Number(videoStyle.line.width) > 0 ? videoStyle.line : null;
    const lineColor = line?.color ?? line?.gradient?.stops[0]?.color ?? "#1f2937";
    if (outline === "none" || (!outline && !line)) {
      return { visible: false, color: lineColor, width: line ? Number(line.width) / 9525 : DEFAULT_IMAGE_BORDER_WIDTH, style: "solid" };
    }
    return outline
      ? { visible: true, color: outline.color, width: outline.widthPx, style: outline.dash }
      : { visible: true, color: lineColor, width: Number(line!.width) / 9525, style: "solid" };
  }
  const width = normalizeImageBorderWidth(record.borderWidth);
  const style = record.borderStyle;
  return {
    visible: style !== "none" && typeof record.borderColor === "string" && (width > 0 || style !== undefined),
    color: normalizeImageBorderColor(record.borderColor, "#1f2937"),
    width: width || DEFAULT_IMAGE_BORDER_WIDTH,
    style: style === "dashed" || style === "dotted" ? style : "solid",
  };
}

/**
 * The block with Video Border set: No Outline writes `borderStyle: "none"`
 * (PowerPoint's explicit `a:ln/a:noFill`), a visible line its colour, width
 * and dash. Returns null when nothing changes.
 */
export function withVideoBorder(
  block: PresentationEmbedBlock,
  payload: VideoBorderPayload,
): PresentationEmbedBlock | null {
  if (!videoBorderEditable(block)) return null;
  const next = { ...block } as Record<string, unknown>;
  for (const key of BORDER_KEYS) delete next[key];
  if (payload.style === "none") {
    next.borderStyle = "none";
  } else {
    const current = videoBorderState(block);
    next.borderStyle = payload.style;
    next.borderColor = normalizeImageBorderColor(payload.color ?? current.color, current.color);
    next.borderWidth = normalizeImageBorderWidth(payload.width ?? current.width) || DEFAULT_IMAGE_BORDER_WIDTH;
  }
  const same = BORDER_KEYS.every((key) => (block as Record<string, unknown>)[key] === next[key]);
  return same ? null : (next as PresentationEmbedBlock);
}

/** The border drawn on the video's frame, centred on its edge as PowerPoint draws a line. */
export function videoBorderCss(block: PresentationEmbedBlock): CSSProperties | null {
  if (block.kind !== "video") return null;
  const state = videoBorderState(block);
  if (!state.visible) return null;
  return {
    outline: `${state.width}px ${state.style} ${state.color}`,
    outlineOffset: `${-state.width / 2}px`,
  };
}

/** Video Format > Video Shape: the gallery kind the video is drawn in. */
export function videoShapeKind(block: PresentationEmbedBlock): ShapeKind {
  const preset = normalizePictureClipShape((block as Record<string, unknown>).clipShape);
  return (preset ? presentationShapeCatalogEntryForPreset(preset)?.kind : null) ?? "rectangle";
}

/**
 * The video drawn in `kind`'s outline (the movie's own `a:prstGeom`, default
 * adjustments, kept in `clipShape` as a picture's Crop to Shape); Rectangle
 * removes it. Null when nothing changes or the kind has no inside.
 */
export function withVideoShape(block: PresentationEmbedBlock, kind: ShapeKind): PresentationEmbedBlock | null {
  if (block.kind !== "video" || (block as Record<string, unknown>).videoStyle !== undefined) return null;
  const preset = kind === "rectangle" ? null : pictureCropShapePreset(kind);
  if (kind !== "rectangle" && !preset) return null;
  const current = normalizePictureClipShape((block as Record<string, unknown>).clipShape);
  if (current === preset) return null;
  const next = { ...block } as Record<string, unknown>;
  delete next.clipShape;
  if (preset) next.clipShape = preset;
  return next as PresentationEmbedBlock;
}

/** The video design fields Video Format > Reset discards (PPT-A155: with its
 * corrections, color, style and the poster the user set). */
const VIDEO_DESIGN_KEYS = [
  "borderColor",
  "borderWidth",
  "borderStyle",
  "pictureEffects",
  "clipShape",
  "videoCorrections",
  "videoRecolor",
  "videoStyle",
  "posterSrc",
] as const;

/** Whether Reset would change anything on this video. */
export function videoHasDesign(block: PresentationEmbedBlock): boolean {
  return block.kind === "video" && VIDEO_DESIGN_KEYS.some((key) => key in (block as Record<string, unknown>));
}

/**
 * Video Format > Adjust > Reset: discards the video's formatting -- its
 * border, effects and shape -- and keeps its media, size and playback
 * settings (Microsoft Support: "Reset Design discards all formatting changes
 * made to this video"). A line the file keeps as its own is not touched.
 * Null when nothing changes.
 */
export function resetVideoDesign(block: PresentationEmbedBlock): PresentationEmbedBlock | null {
  if (!videoHasDesign(block)) return null;
  const next = { ...block } as Record<string, unknown>;
  for (const key of VIDEO_DESIGN_KEYS) {
    if (key.startsWith("border") && !videoBorderEditable(block)) continue;
    if ((key === "videoCorrections" || key === "videoRecolor") && !videoAdjustEditable(block)) continue;
    delete next[key];
  }
  const same = VIDEO_DESIGN_KEYS.every((key) => (block as Record<string, unknown>)[key] === next[key]);
  return same ? null : (next as PresentationEmbedBlock);
}

// --- Video Format > Adjust (Corrections, Color, Poster Frame) and Video Styles (PPT-A155) ----

/**
 * Corrections: whole signed percentages (Office's -100..100), never both
 * zero on the page. Color: one Recolor preset (`office-recolor-presets`).
 * Washout is Corrections +70 / -70, so it and Corrections replace each
 * other. The file side is the server's `video_adjust`.
 */
export interface VideoCorrections {
  readonly brightness: number;
  readonly contrast: number;
}

/** Whether the movie's Corrections / Color can be restated (the file's own blip is kept). */
export function videoAdjustEditable(block: PresentationEmbedBlock | null | undefined): boolean {
  if (!block || block.kind !== "video") return false;
  const source = block.source as Record<string, unknown> | undefined;
  return source?.videoAdjustDisposition !== "preserve-only";
}

const wholePercent = (value: unknown): number =>
  typeof value === "number" && Number.isInteger(value) && value >= -100 && value <= 100 ? value : 0;

export function videoCorrections(block: PresentationEmbedBlock): VideoCorrections {
  const raw = (block as Record<string, unknown>).videoCorrections as Record<string, unknown> | undefined;
  return { brightness: wholePercent(raw?.brightness), contrast: wholePercent(raw?.contrast) };
}

export function videoRecolor(block: PresentationEmbedBlock): OfficeRecolorPresetId | null {
  const value = (block as Record<string, unknown>).videoRecolor;
  return isOfficeRecolorPreset(value) ? value : null;
}

function sameRecord(left: Record<string, unknown>, right: Record<string, unknown>, keys: readonly string[]) {
  return keys.every((key) => JSON.stringify(left[key]) === JSON.stringify(right[key]));
}

/** Video Corrections: a 5x5 tile or the pane's fields. Null when nothing changes. */
export function withVideoCorrections(
  block: PresentationEmbedBlock,
  brightness: number,
  contrast: number,
): PresentationEmbedBlock | null {
  if (!videoAdjustEditable(block)) return null;
  const next = { ...block } as Record<string, unknown>;
  const b = Math.round(Math.max(-100, Math.min(100, brightness)));
  const c = Math.round(Math.max(-100, Math.min(100, contrast)));
  delete next.videoCorrections;
  if (b || c) next.videoCorrections = { brightness: b, contrast: c };
  // Washout is this same luminance: a correction takes its place.
  if ((b || c) && next.videoRecolor === "washout") delete next.videoRecolor;
  return sameRecord(block as Record<string, unknown>, next, ["videoCorrections", "videoRecolor"])
    ? null
    : (next as PresentationEmbedBlock);
}

/** Video Color > Recolor: a preset, or No Recolor (null). Null when nothing changes. */
export function withVideoRecolor(
  block: PresentationEmbedBlock,
  preset: OfficeRecolorPresetId | null,
): PresentationEmbedBlock | null {
  if (!videoAdjustEditable(block)) return null;
  const next = { ...block } as Record<string, unknown>;
  delete next.videoRecolor;
  if (preset) next.videoRecolor = preset;
  if (preset === "washout") delete next.videoCorrections;
  return sameRecord(block as Record<string, unknown>, next, ["videoCorrections", "videoRecolor"])
    ? null
    : (next as PresentationEmbedBlock);
}

/** The SVG filter steps the movie is drawn through (canvas and slide show). */
export function videoAdjustFilterSteps(
  block: PresentationEmbedBlock,
  theme?: OfficeThemeGridColors | null,
): OfficeAdjustFilterStep[] {
  if (block.kind !== "video") return [];
  const corrections = videoCorrections(block);
  return officeAdjustFilterSteps({ recolor: videoRecolor(block), ...corrections }, theme);
}

/** Video Styles: the movie's style, when one of the 41 is applied. */
export function videoStyleId(block: PresentationEmbedBlock): OfficeVideoStyleId | null {
  const value = (block as Record<string, unknown>).videoStyle;
  return officeVideoStyle(value) ? (value as OfficeVideoStyleId) : null;
}

/**
 * A video style is a whole recipe -- geometry, outline, effects, 3-D: while
 * one is applied the movie's Video Shape and Video Border are the style's
 * (as Picture Border and Crop to Shape are a picture style's).
 */
export function videoStyleOwnsFrame(block: PresentationEmbedBlock): boolean {
  return videoStyleId(block) !== null;
}

/**
 * Video Styles: apply one (the recipe replaces the movie's border, shape and
 * effects, as PowerPoint does), or none. Null when nothing changes.
 */
export function withVideoStyle(
  block: PresentationEmbedBlock,
  styleId: OfficeVideoStyleId | null,
): PresentationEmbedBlock | null {
  if (block.kind !== "video") return null;
  if (styleId !== null && !officeVideoStyle(styleId)) return null;
  const next = { ...block } as Record<string, unknown>;
  delete next.videoStyle;
  if (styleId) {
    next.videoStyle = styleId;
    for (const key of ["borderColor", "borderWidth", "borderStyle", "clipShape", "pictureEffects"]) delete next[key];
  }
  return sameRecord(block as Record<string, unknown>, next, [
    "videoStyle",
    "borderColor",
    "borderWidth",
    "borderStyle",
    "clipShape",
    "pictureEffects",
  ])
    ? null
    : (next as PresentationEmbedBlock);
}

/**
 * Poster Frame: the picture the video shows until it plays -- Current Frame
 * (a canvas grab of the playing frame) or Image from File, as an inline PNG
 * or JPEG of at most 1 MiB of text (`posterSrc`); Reset (null) returns to the
 * poster the video had. Null when nothing changes or the source is not one.
 */
export function withVideoPoster(block: PresentationEmbedBlock, posterSrc: string | null): PresentationEmbedBlock | null {
  if (block.kind !== "video") return null;
  if (posterSrc !== null && !isVideoPosterSource(posterSrc)) return null;
  const current = (block as Record<string, unknown>).posterSrc;
  if ((current ?? null) === posterSrc) return null;
  const next = { ...block } as Record<string, unknown>;
  delete next.posterSrc;
  if (posterSrc) next.posterSrc = posterSrc;
  return next as PresentationEmbedBlock;
}

export { isVideoPosterSource, MAX_VIDEO_POSTER_SRC_LENGTH };

export function videoPosterSource(block: PresentationEmbedBlock): string | null {
  const value = (block as Record<string, unknown>).posterSrc;
  return isVideoPosterSource(value) ? value : null;
}
