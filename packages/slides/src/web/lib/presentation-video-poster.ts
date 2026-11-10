/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Video Format > Poster Frame > Current Frame, the browser half (PPT-A155):
 * the frame the video shows at its player's position, grabbed through a
 * canvas as a PNG (or, when that is over the page's 1 MiB string bound, a
 * smaller JPEG) data URL the page keeps as `posterSrc`. The server writes it
 * as the movie's poster image; nothing here knows the file format.
 *
 * A frame from another origin taints the canvas (the browser refuses its
 * pixels); the grab then retries through a CORS-mode copy of the same source
 * at the same time and, when the source does not allow that either, says so.
 */

import { fitOfficePictureSource } from "@mrnote/core/web/lib/office-picture-envelope.ts";
import { isVideoPosterSource, MAX_VIDEO_POSTER_SRC_LENGTH } from "@mrnote/core/web/lib/presentation-video-poster-source.ts";

export type VideoFrameCaptureFailure = "not-ready" | "cross-origin" | "too-large";

export class VideoFrameCaptureError extends Error {
  readonly code: VideoFrameCaptureFailure;

  constructor(code: VideoFrameCaptureFailure) {
    super(`presentation_video_frame_${code}`);
    this.name = "VideoFrameCaptureError";
    this.code = code;
  }
}

/** The longest side a poster keeps (a full-HD frame is reduced to this). */
export const VIDEO_POSTER_MAX_SIDE = 1920;

function whenFrameReady(video: HTMLVideoElement, timeoutMs: number): Promise<void> {
  if (video.readyState >= 2 && !video.seeking) return Promise.resolve();
  return new Promise((resolve, reject) => {
    const done = () => {
      cleanup();
      resolve();
    };
    const timer = setTimeout(() => {
      cleanup();
      reject(new VideoFrameCaptureError("not-ready"));
    }, timeoutMs);
    const cleanup = () => {
      clearTimeout(timer);
      video.removeEventListener("loadeddata", done);
      video.removeEventListener("seeked", done);
    };
    video.addEventListener("loadeddata", done);
    video.addEventListener("seeked", done);
  });
}

function drawFrame(video: HTMLVideoElement): string {
  const width = video.videoWidth;
  const height = video.videoHeight;
  if (!(width > 0) || !(height > 0)) throw new VideoFrameCaptureError("not-ready");
  const scale = Math.min(1, VIDEO_POSTER_MAX_SIDE / Math.max(width, height));
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(width * scale));
  canvas.height = Math.max(1, Math.round(height * scale));
  const context = canvas.getContext("2d");
  if (!context) throw new VideoFrameCaptureError("not-ready");
  context.drawImage(video, 0, 0, canvas.width, canvas.height);
  try {
    const png = canvas.toDataURL("image/png");
    if (png.length <= MAX_VIDEO_POSTER_SRC_LENGTH) return png;
    return canvas.toDataURL("image/jpeg", 0.9);
  } catch (error) {
    if (error instanceof DOMException && error.name === "SecurityError") {
      throw new VideoFrameCaptureError("cross-origin");
    }
    throw error;
  }
}

async function corsCopyFrame(video: HTMLVideoElement): Promise<string> {
  const source = video.currentSrc || video.src;
  if (!source || video.crossOrigin === "anonymous") throw new VideoFrameCaptureError("cross-origin");
  const copy = document.createElement("video");
  copy.crossOrigin = "anonymous";
  copy.muted = true;
  copy.preload = "auto";
  copy.src = source;
  try {
    await whenFrameReady(copy, 8_000);
    copy.currentTime = video.currentTime;
    await whenFrameReady(copy, 8_000);
    return drawFrame(copy);
  } catch (error) {
    if (error instanceof VideoFrameCaptureError && error.code !== "not-ready") throw error;
    throw new VideoFrameCaptureError("cross-origin");
  } finally {
    copy.removeAttribute("src");
    copy.load();
  }
}

/** The frame at the player's current position, as a poster source. */
export async function capturePresentationVideoFrame(video: HTMLVideoElement): Promise<string> {
  // A player that only loaded its metadata has no frame yet: seeking to where
  // it stands makes the browser decode that frame.
  if (video.readyState < 2 && video.readyState >= 1) video.currentTime = video.currentTime;
  await whenFrameReady(video, 5_000);
  let dataUrl: string;
  try {
    dataUrl = drawFrame(video);
  } catch (error) {
    if (!(error instanceof VideoFrameCaptureError) || error.code !== "cross-origin") throw error;
    dataUrl = await corsCopyFrame(video);
  }
  const fitted = await fitOfficePictureSource(dataUrl, undefined, MAX_VIDEO_POSTER_SRC_LENGTH);
  if (!fitted || !isVideoPosterSource(fitted)) throw new VideoFrameCaptureError("too-large");
  return fitted;
}
