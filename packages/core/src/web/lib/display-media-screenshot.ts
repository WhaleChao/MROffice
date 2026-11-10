/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export type DisplayMediaScreenshotResult =
  | { readonly kind: "captured"; readonly file: File }
  | { readonly kind: "cancelled" | "failed" | "unsupported" };

type DisplayMediaScreenshotOptions = {
  readonly maxBytes?: number;
  readonly maxEdge?: number;
  readonly maxPixels?: number;
  readonly now?: () => number;
  readonly createCanvas?: () => HTMLCanvasElement;
  readonly createVideo?: () => HTMLVideoElement;
  readonly getDisplayMedia?: (
    constraints: DisplayMediaStreamOptions,
  ) => Promise<MediaStream>;
};

const DEFAULT_MAX_EDGE = 4_096;
const DEFAULT_MAX_PIXELS = 16_777_216;
const DEFAULT_MAX_BYTES = 8 * 1024 * 1024;

function canvasBlob(
  canvas: HTMLCanvasElement,
  type: "image/jpeg" | "image/png",
  quality?: number,
): Promise<Blob | null> {
  return new Promise((resolve) => canvas.toBlob(resolve, type, quality));
}

function boundedCaptureSize(
  width: number,
  height: number,
  maxEdge: number,
  maxPixels: number,
): { width: number; height: number } {
  const safeWidth = Math.max(1, Math.floor(width));
  const safeHeight = Math.max(1, Math.floor(height));
  const scale = Math.min(
    1,
    maxEdge / safeWidth,
    maxEdge / safeHeight,
    Math.sqrt(maxPixels / (safeWidth * safeHeight)),
  );
  return {
    width: Math.max(1, Math.floor(safeWidth * scale)),
    height: Math.max(1, Math.floor(safeHeight * scale)),
  };
}

async function encodeBoundedScreenshot(
  video: HTMLVideoElement,
  canvas: HTMLCanvasElement,
  maxBytes: number,
  maxEdge: number,
  maxPixels: number,
): Promise<{ blob: Blob; extension: "jpg" | "png" } | null> {
  const initial = boundedCaptureSize(
    video.videoWidth || 1_280,
    video.videoHeight || 720,
    maxEdge,
    maxPixels,
  );
  let width = initial.width;
  let height = initial.height;
  for (let attempt = 0; attempt < 5; attempt += 1) {
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext("2d");
    if (!context) return null;
    context.drawImage(video, 0, 0, width, height);
    const png = await canvasBlob(canvas, "image/png");
    if (png && png.size > 0 && png.size <= maxBytes) {
      return { blob: png, extension: "png" };
    }
    for (const quality of [0.9, 0.78, 0.64] as const) {
      const jpeg = await canvasBlob(canvas, "image/jpeg", quality);
      if (jpeg && jpeg.size > 0 && jpeg.size <= maxBytes) {
        return { blob: jpeg, extension: "jpg" };
      }
    }
    width = Math.max(1, Math.floor(width * 0.75));
    height = Math.max(1, Math.floor(height * 0.75));
  }
  return null;
}

const DISPLAY_MEDIA_SOURCE_FAILURES = new Set([
  "NotReadableError",
  "NotFoundError",
  "InvalidStateError",
  "OverconstrainedError",
]);

/**
 * Capture one user-selected screen/window/tab frame as a bounded image file.
 * The browser owns the privacy chooser; no capture begins without that explicit
 * user gesture. Every acquired track is stopped on success, denial, or failure.
 */
export async function captureDisplayMediaScreenshot(
  options: DisplayMediaScreenshotOptions = {},
): Promise<DisplayMediaScreenshotResult> {
  const getDisplayMedia =
    options.getDisplayMedia ??
    navigator.mediaDevices?.getDisplayMedia?.bind(navigator.mediaDevices);
  if (!getDisplayMedia) return { kind: "unsupported" };
  let stream: MediaStream | null = null;
  let acquired = false;
  let video: HTMLVideoElement | null = null;
  try {
    stream = await getDisplayMedia({ video: true, audio: false });
    acquired = true;
    video = options.createVideo?.() ?? document.createElement("video");
    video.srcObject = stream;
    video.muted = true;
    video.playsInline = true;
    await video.play();
    const canvas = options.createCanvas?.() ?? document.createElement("canvas");
    const encoded = await encodeBoundedScreenshot(
      video,
      canvas,
      Math.max(1, options.maxBytes ?? DEFAULT_MAX_BYTES),
      Math.max(1, options.maxEdge ?? DEFAULT_MAX_EDGE),
      Math.max(1, options.maxPixels ?? DEFAULT_MAX_PIXELS),
    );
    if (!encoded) return { kind: "failed" };
    const type = encoded.extension === "png" ? "image/png" : "image/jpeg";
    return {
      kind: "captured",
      file: new File(
        [encoded.blob],
        `screenshot-${(options.now ?? Date.now)()}.${encoded.extension}`,
        { type },
      ),
    };
  } catch (error) {
    // The chooser dismissed (NotAllowedError/AbortError) is the user's
    // choice and says nothing. A source the system would not start (screen
    // recording refused to the browser, no screen, a capture already
    // running) is a failure the user is told of: read as a cancel, the
    // command did nothing at all.
    // A DOMException is not an Error in every runtime: read its name.
    const name = String((error as { name?: unknown } | null)?.name ?? "");
    return {
      kind:
        acquired || DISPLAY_MEDIA_SOURCE_FAILURES.has(name) ? "failed" : "cancelled",
    };
  } finally {
    if (video) video.srcObject = null;
    for (const track of stream?.getTracks() ?? []) track.stop();
  }
}

/** A region of a captured frame, in fractions of its width and height. */
export type ScreenClippingRect = {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
};

/** The smallest region Screen Clipping takes, in pixels of the shown frame:
 * a click without a drag clips nothing (Office inserts nothing either). */
export const SCREEN_CLIPPING_MIN_PIXELS = 4;

/**
 * Insert > Screenshot > Screen Clipping: the rectangle the user dragged from
 * `start` to `end` over a frame shown at `bounds` (both corners in the same
 * client coordinates), clamped to the frame, as fractions of it. Null when
 * the drag is smaller than `SCREEN_CLIPPING_MIN_PIXELS` either way.
 */
export function screenClippingRect(
  start: { readonly x: number; readonly y: number },
  end: { readonly x: number; readonly y: number },
  bounds: { readonly left: number; readonly top: number; readonly width: number; readonly height: number },
): ScreenClippingRect | null {
  if (!(bounds.width > 0) || !(bounds.height > 0)) return null;
  const clampX = (value: number) => Math.min(bounds.left + bounds.width, Math.max(bounds.left, value));
  const clampY = (value: number) => Math.min(bounds.top + bounds.height, Math.max(bounds.top, value));
  const left = Math.min(clampX(start.x), clampX(end.x));
  const right = Math.max(clampX(start.x), clampX(end.x));
  const top = Math.min(clampY(start.y), clampY(end.y));
  const bottom = Math.max(clampY(start.y), clampY(end.y));
  if (right - left < SCREEN_CLIPPING_MIN_PIXELS || bottom - top < SCREEN_CLIPPING_MIN_PIXELS) {
    return null;
  }
  return {
    x: (left - bounds.left) / bounds.width,
    y: (top - bounds.top) / bounds.height,
    width: (right - left) / bounds.width,
    height: (bottom - top) / bounds.height,
  };
}

/** The pixel box of `rect` in an image of `width` x `height`, at least 1 x 1. */
export function screenClippingPixels(
  rect: ScreenClippingRect,
  width: number,
  height: number,
): { sx: number; sy: number; sw: number; sh: number } {
  const sx = Math.min(width - 1, Math.max(0, Math.round(rect.x * width)));
  const sy = Math.min(height - 1, Math.max(0, Math.round(rect.y * height)));
  return {
    sx,
    sy,
    sw: Math.max(1, Math.min(width - sx, Math.round(rect.width * width))),
    sh: Math.max(1, Math.min(height - sy, Math.round(rect.height * height))),
  };
}

/** Cut `rect` out of a captured frame as a new PNG file (the picture Screen
 * Clipping inserts). Null when the browser cannot decode or encode it. */
export async function clipScreenshotFile(
  file: File,
  rect: ScreenClippingRect,
  options: { readonly now?: () => number } = {},
): Promise<File | null> {
  let bitmap: ImageBitmap | null = null;
  try {
    bitmap = await createImageBitmap(file);
    const box = screenClippingPixels(rect, bitmap.width, bitmap.height);
    const canvas = document.createElement("canvas");
    canvas.width = box.sw;
    canvas.height = box.sh;
    const context = canvas.getContext("2d");
    if (!context) return null;
    context.drawImage(bitmap, box.sx, box.sy, box.sw, box.sh, 0, 0, box.sw, box.sh);
    const blob = await canvasBlob(canvas, "image/png");
    if (!blob || blob.size === 0) return null;
    return new File([blob], `screen-clipping-${(options.now ?? Date.now)()}.png`, {
      type: "image/png",
    });
  } catch {
    return null;
  } finally {
    bitmap?.close();
  }
}
